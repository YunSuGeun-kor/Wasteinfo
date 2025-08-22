# app.py
import os
import re
import unicodedata
import traceback
from typing import Optional, List, Tuple, Dict, Any

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

# =========================================
# 설정
# =========================================
APP_TITLE = "사내 폐기물 처리방법 조회"
DEFAULT_XLSX = r"C:/Users/cf100/Desktop/.streamlit/폐기물처리방법.xlsx"  # 기본 경로 (사이드바/환경변수로 변경 가능)

# =========================================
# 유틸
# =========================================
def normalize_korean(s: str) -> str:
    """NFKC 정규화 + 공백 정리(간단)"""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

# =========================================
# 데이터 로딩
# =========================================
def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # '처리 방법 '처럼 끝 공백이 있을 수 있어 strip
    df.columns = [c.strip() for c in df.columns]
    return df

@st.cache_data(show_spinner=False)
def load_data(xlsx_path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    시트:
      - '광양소 폐기물처리방법' (메인)
      - '폐기물관리법_시행규칙_별표5' (참고)
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(xlsx_path)

    df_main = pd.read_excel(xlsx_path, sheet_name="광양소 폐기물처리방법", engine="openpyxl")
    df_ref  = pd.read_excel(xlsx_path, sheet_name="폐기물관리법_시행규칙_별표5", engine="openpyxl")

    df_main = _normalize_columns(df_main)
    df_ref  = _normalize_columns(df_ref)

    # 필수 컬럼 체크
    for req in ["폐기물 종류", "처리 방법"]:
        if req not in df_main.columns:
            raise ValueError(f"메인 시트에 '{req}' 컬럼이 없습니다.")
    for req in ["구분", "폐기물 종류", "처리 방법"]:
        if req not in df_ref.columns:
            raise ValueError(f"참고 시트에 '{req}' 컬럼이 없습니다.")

    # 문자열화/트림
    df_main["폐기물 종류"] = df_main["폐기물 종류"].astype(str).str.strip()
    df_main["처리 방법"]   = df_main["처리 방법"].astype(str).str.strip()
    df_ref["구분"]        = df_ref["구분"].astype(str).str.strip()
    df_ref["폐기물 종류"]  = df_ref["폐기물 종류"].astype(str).str.strip()
    df_ref["처리 방법"]    = df_ref["처리 방법"].astype(str).str.strip()

    return df_main, df_ref

# =========================================
# 매칭/검색
# =========================================
def normalize_query(name: str, phase: Optional[str], material: str, openai_client=None) -> List[str]:
    """
    입력을 표준화된 후보 검색어 리스트로 변환.
    - OpenAI가 있으면 동의어/오타 보정 JSON 배열을 받아 확장(선택).
    - 없으면 간단 규칙 기반 확장.
    """
    base_terms = []
    for s in [name, material, phase or ""]:
        s = (s or "").strip()
        if s:
            base_terms.append(normalize_korean(s))

    candidates = list(dict.fromkeys([t for t in base_terms if t]))  # 중복 제거

    if openai_client:
        try:
            sys = "너는 사내 폐기물 용어 표준화 도우미다. 사용자 입력을 표준 용어 후보 JSON 배열로만 반환하라."
            user = {"name": name, "phase": phase, "material": material}
            rsp = openai_client.responses.create(
                model="gpt-4o-mini",
                input=[
                    {"role": "system", "content": sys},
                    {"role": "user", "content": str(user)}
                ],
                temperature=0
            )
            text = rsp.output_text.strip()
            if text.startswith("[") and text.endswith("]"):
                import json
                arr = json.loads(text)
                arr = [normalize_korean(str(x)) for x in arr if str(x).strip()]
                candidates = list(dict.fromkeys(candidates + arr))
        except Exception:
            # 실패 시 무시하고 퍼지 매칭만 사용
            pass

    return candidates or [normalize_korean(name)]

def _score_series(series: pd.Series, query: str) -> List[Tuple[int, float]]:
    """series(폐기물 종류) 각 항목과 query 유사도를 계산 → (index, score) 리스트"""
    scores = []
    qn = normalize_korean(query)
    for idx, val in series.items():
        s = normalize_korean(val)
        score = fuzz.WRatio(qn, s)  # 종합 가중 유사도
        scores.append((idx, float(score)))
    return scores

def search_best(df_main: pd.DataFrame, query_terms: List[str], threshold: int = 60) -> Tuple[Optional[pd.Series], Dict[str, Any]]:
    """
    1) 정확/부분 일치 우선
    2) 퍼지 스코어 기반 최상위 1건
    """
    col_name = "폐기물 종류"
    col_method = "처리 방법"
    debug: Dict[str, Any] = {"match_type": "", "score": 0.0, "candidates": []}

    # 1) 정확/부분 일치
    for q in query_terms:
        # 정확 일치
        exact = df_main[df_main[col_name].str.strip().str.lower() == q.strip().lower()]
        if len(exact) > 0:
            row = exact.iloc[0]
            debug["match_type"] = "정확 일치"
            debug["score"] = 100.0
            return row, debug
        # 부분 일치
        part = df_main[df_main[col_name].str.contains(q, case=False, na=False)]
        if len(part) > 0:
            row = part.iloc[0]
            debug["match_type"] = "부분 일치"
            debug["score"] = 95.0
            return row, debug

    # 2) 퍼지 매칭
    all_scores: List[Tuple[int, float]] = []
    for q in query_terms:
        all_scores.extend(_score_series(df_main[col_name], q))
    all_scores.sort(key=lambda x: -x[1])
    top = all_scores[:10]

    debug["candidates"] = []
    seen = set()
    best_row = None
    best_score = 0.0

    for idx, sc in top:
        if idx in seen:
            continue
        seen.add(idx)
        row = df_main.loc[idx]
        debug["candidates"].append({
            "폐기물 종류": row[col_name],
            "처리 방법": row[col_method],
            "score": sc
        })
        if sc > best_score:
            best_row = row
            best_score = sc

    if best_row is not None and best_score >= threshold:
        debug["match_type"] = "퍼지 매칭"
        debug["score"] = best_score
        return best_row, debug

    return None, debug

def find_refs(df_ref: pd.DataFrame, keyword: str, topk: int = 3) -> list[dict]:
    """참고(법령) 시트에서 유사 항목 상위 topk 반환"""
    col_name = "폐기물 종류"
    col_method = "처리 방법"
    col_group = "구분"

    scores = []
    key_norm = normalize_korean(keyword)
    for idx, val in df_ref[col_name].items():
        s = normalize_korean(val)
        score = fuzz.WRatio(key_norm, s)
        scores.append((idx, float(score)))

    scores.sort(key=lambda x: -x[1])
    out = []
    for idx, sc in scores[:topk]:
        row = df_ref.loc[idx]
        out.append({
            "구분": row[col_group],
            "폐기물 종류": row[col_name],
            "처리 방법": row[col_method],
            "score": sc
        })
    return out

# =========================================
# Streamlit App
# =========================================
st.set_page_config(page_title=APP_TITLE, page_icon="♻️", layout="wide")

# 세션 상태
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False
    st.session_state.df_main = None
    st.session_state.df_ref = None
    st.session_state.xlsx_path = None
if "OPENAI_API_KEY" not in st.session_state:
    # 환경변수에서 초기값 채우기(있으면)
    env_key = os.getenv("OPENAI_API_KEY")
    if env_key:
        st.session_state.OPENAI_API_KEY = env_key

# 경로 결정 우선순위: 사이드바 입력 > 환경변수(DATA_XLSX) > 기본값
def effective_path(default_rel: str = DEFAULT_XLSX):
    sb = st.session_state.get("sidebar_path")
    if sb and sb.strip():
        return sb.strip()
    env = os.getenv("DATA_XLSX")
    if env and env.strip():
        return env.strip()
    return default_rel

def load_app_data(xlsx_path: str):
    """데이터 로딩 + 상태 저장"""
    try:
        df_main, df_ref = load_data(xlsx_path)
        st.session_state.df_main = df_main
        st.session_state.df_ref = df_ref
        st.session_state.xlsx_path = xlsx_path
        st.session_state.data_loaded = True
        return True
    except FileNotFoundError:
        st.error(f"📁 엑셀 파일을 찾을 수 없습니다: `{xlsx_path}`\n경로를 확인하세요.")
        return False
    except Exception as e:
        st.error(f"❌ 데이터 로딩 중 오류: {str(e)}")
        with st.expander("상세 오류 정보"):
            st.code(traceback.format_exc())
        return False

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("🔧 시스템 관리")

    # 엑셀 경로
    st.caption("엑셀 파일 경로를 지정하세요. (상대/절대 경로 가능)")
    sidebar_path = st.text_input(
        "엑셀 파일 경로",
        value=st.session_state.get("sidebar_path", effective_path()),
        help="예: ./data/폐기물처리방법.xlsx 또는 C:/work/폐기물처리방법.xlsx"
    )
    st.session_state.sidebar_path = sidebar_path

    if st.button("📊 데이터 불러오기/새로고침", use_container_width=True):
        st.session_state.data_loaded = False
        _ = load_app_data(effective_path())
        st.rerun()

    # 데이터 상태
    if st.session_state.data_loaded:
        st.success("✅ 데이터 로드됨")
        st.info(f"📂 경로: `{st.session_state.xlsx_path}`")
        if st.session_state.df_main is not None:
            st.caption(f"📋 메인 데이터: {len(st.session_state.df_main)}개 항목")
        if st.session_state.df_ref is not None:
            st.caption(f"📖 참고 데이터: {len(st.session_state.df_ref)}개 항목")
    else:
        st.warning("⚠️ 데이터 미로드")

    # OpenAI API Key 입력(웹에서)
    st.subheader("🤖 OpenAI 연결 설정")
    api_key_input = st.text_input(
        "OpenAI API Key",
        type="password",
        placeholder="sk-로 시작하는 API Key 입력"
    )
    if api_key_input:
        st.session_state["OPENAI_API_KEY"] = api_key_input

    # 상태 표시
    if st.session_state.get("OPENAI_API_KEY"):
        st.success("🔑 OpenAI 연결됨")
    else:
        st.info("💡 OpenAI 미연결 (퍼지 매칭 사용)")

# ---------------- Main ----------------
st.title(f"♻️ {APP_TITLE}")
st.markdown("---")

# 첫 로드 시 자동 시도
if not st.session_state.data_loaded:
    with st.spinner("📊 데이터를 로딩중입니다..."):
        load_app_data(effective_path())

if st.session_state.data_loaded:
    st.subheader("🔍 폐기물 정보 입력")
    col1, col2, col3 = st.columns([2, 1, 2])

    with col1:
        waste_name = st.text_input(
            "폐기물명 *",
            placeholder="예: 폐유, 폐페인트 슬러지",
            help="처리하고자 하는 폐기물의 이름을 입력하세요"
        )
    with col2:
        phase = st.selectbox(
            "성상",
            ["선택안함", "고체", "액체"],
            help="폐기물의 물리적 상태를 선택하세요"
        )
    with col3:
        material = st.text_input(
            "재질",
            placeholder="예: PET, 고무, 유기용제 함유",
            help="폐기물의 재질이나 구성 성분을 입력하세요"
        )

    st.markdown("---")
    col_center = st.columns([1, 2, 1])[1]
    with col_center:
        search_button = st.button("🔍 처리방법 조회", type="primary", use_container_width=True)

    if search_button:
        if not waste_name.strip():
            st.warning("⚠️ 폐기물명을 입력해주세요.")
        else:
            with st.spinner("🔍 검색중입니다..."):
                try:
                    phase_input = None if phase == "선택안함" else phase

                    # 선택적 OpenAI: 세션 상태의 키 우선 사용
                    openai_client = None
                    api_key = st.session_state.get("OPENAI_API_KEY")
                    if api_key:
                        try:
                            from openai import OpenAI
                            openai_client = OpenAI(api_key=api_key)
                        except Exception as e:
                            st.warning(f"⚠️ OpenAI 초기화 실패: {e}. 퍼지 매칭만 사용합니다.")

                    # Normalize + 매칭
                    query_terms = normalize_query(waste_name, phase_input, material, openai_client)
                    best_row, debug_info = search_best(st.session_state.df_main, query_terms)

                    if best_row is not None:
                        st.subheader("✅ 처리 방법 (사내 기준)")
                        with st.container():
                            st.success(f"**처리 방법**: {best_row['처리 방법']}")
                            col_details1, col_details2 = st.columns(2)
                            with col_details1:
                                st.info(f"**매칭된 폐기물**: {best_row['폐기물 종류']}")
                            with col_details2:
                                match_type = debug_info.get("match_type", "알 수 없음")
                                score = debug_info.get("score", 0)
                                st.info(f"**매칭 방식**: {match_type} (유사도: {score:.1f}%)")

                        # 법령 참고
                        refs = find_refs(st.session_state.df_ref, best_row["폐기물 종류"])
                        if refs:
                            st.markdown("---")
                            st.subheader("📖 법령 참고 (폐기물관리법 시행규칙 별표5)")
                            with st.expander("참고 정보 보기", expanded=False):
                                for i, ref in enumerate(refs[:3], 1):
                                    st.markdown(f"**{i}. {ref['구분']}**")
                                    st.markdown(f"- 폐기물 종류: {ref['폐기물 종류']}")
                                    st.markdown(f"- 처리 방법: {ref['처리 방법']}")
                                    if i < len(refs[:3]):
                                        st.markdown("---")
                                st.warning("⚠️ **주의**: 법령 정보는 참고이며, 실제 업무 적용은 사내 기준을 우선합니다.")

                        # 디버그
                        if debug_info.get("candidates"):
                            with st.expander("🔍 검색 상세 정보", expanded=False):
                                st.json({
                                    "입력_정보": {
                                        "폐기물명": waste_name,
                                        "성상": phase_input,
                                        "재질": material,
                                        "정규화된_검색어": query_terms
                                    },
                                    "매칭_결과": debug_info
                                })
                    else:
                        st.error("❌ 일치하는 항목을 찾지 못했습니다.")
                        if debug_info.get("candidates"):
                            st.subheader("💡 유사한 항목들")
                            for i, c in enumerate(debug_info["candidates"][:5], 1):
                                st.markdown(f"**{i}. {c['폐기물 종류']}** (유사도: {c['score']:.1f}%)")
                                st.markdown(f"   처리 방법: {c['처리 방법']}")
                        st.info("📞 정확한 처리방법은 환경자원그룹(061-790-8526)에 문의하시기 바랍니다.")
                except Exception as e:
                    st.error(f"❌ 검색 중 오류가 발생했습니다: {str(e)}")
                    with st.expander("상세 오류 정보"):
                        st.code(traceback.format_exc())
else:
    st.error("❌ 데이터를 로드할 수 없습니다. 사이드바의 '데이터 불러오기/새로고침'을 클릭하거나 경로를 확인하세요.")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>"
    "💡 시스템 문의: 환경자원그룹(061-790-8526) | "
    "📧 실제 업무 적용은 사내 기준을 우선하세요"
    "</div>",
    unsafe_allow_html=True
)
