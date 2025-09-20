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
DEFAULT_XLSX = "wasteinfo.xlsx"  # 리포지토리 루트에 배치

# =========================================
# 유틸
# =========================================
def normalize_korean(s: str) -> str:
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def get_dept_from_row(row: pd.Series) -> str:
    # '부서' 관련 컬럼 탐색(예: '부서', '담당 부서', '처리부서')
    for c in row.index:
        if "부서" in str(c):
            v = str(row[c]).strip()
            if v and v != "nan":
                return v
    return ""

# =========================================
# 데이터 로딩
# =========================================
def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    return df

@st.cache_data(show_spinner=False)
def load_data(xlsx_path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
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

    # 문자열 정리
    for c in ["폐기물 종류", "처리 방법"]:
        df_main[c] = df_main[c].astype(str).str.strip()
    for c in ["구분", "폐기물 종류", "처리 방법"]:
        df_ref[c] = df_ref[c].astype(str).str.strip()

    return df_main, df_ref

# =========================================
# 매칭/검색
# =========================================
def normalize_query(name: str, phase: Optional[str], material: str, openai_client=None) -> List[str]:
    base_terms = []
    for s in [name, material, phase or ""]:
        s = (s or "").strip()
        if s:
            base_terms.append(normalize_korean(s))
    candidates = list(dict.fromkeys([t for t in base_terms if t]))

    if openai_client:
        try:
            sys_msg = "너는 사내 폐기물 용어 표준화 도우미다. 사용자 입력을 표준 용어 후보 JSON 배열로만 반환하라."
            user = {"name": name, "phase": phase, "material": material}
            rsp = openai_client.responses.create(
                model="gpt-4o-mini",
                input=[
                    {"role": "system", "content": sys_msg},
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
            pass
    return candidates or [normalize_korean(name)]

def _score_series(series: pd.Series, query: str) -> List[Tuple[int, float]]:
    scores = []
    qn = normalize_korean(query)
    for idx, val in series.items():
        s = normalize_korean(val)
        score = fuzz.WRatio(qn, s)
        scores.append((idx, float(score)))
    return scores

def search_best(df_main: pd.DataFrame, query_terms: List[str], threshold: int = 60) -> Tuple[Optional[pd.Series], Dict[str, Any]]:
    col_name, col_method = "폐기물 종류", "처리 방법"
    debug: Dict[str, Any] = {"match_type": "", "score": 0.0, "candidates": []}

    # 1) 정확/부분 일치
    for q in query_terms:
        ql = q.strip().lower()
        exact = df_main[df_main[col_name].str.strip().str.lower() == ql]
        if len(exact) > 0:
            row = exact.iloc[0]
            debug["match_type"] = "정확 일치"
            debug["score"] = 100.0
            return row, debug

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

    best_row, best_score = None, 0.0
    seen = set()
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
            best_row, best_score = row, sc

    if best_row is not None and best_score >= threshold:
        debug["match_type"] = "퍼지 매칭"
        debug["score"] = best_score
        return best_row, debug

    return None, debug

def find_best_ref_one(df_ref: pd.DataFrame, keyword: str) -> Optional[dict]:
    """'구분'에 '처리기준' 포함 행을 우선 필터 → 유사도 최고 1건 반환. 없으면 전체 중 1건."""
    col_name, col_method, col_group = "폐기물 종류", "처리 방법", "구분"
    key_norm = normalize_korean(keyword)

    def pick_best(sub: pd.DataFrame) -> Optional[dict]:
        if sub.empty:
            return None
        scores = []
        for idx, val in sub[col_name].items():
            s = normalize_korean(val)
            sc = fuzz.WRatio(key_norm, s)
            scores.append((idx, float(sc)))
        scores.sort(key=lambda x: -x[1])
        idx, sc = scores[0]
        row = sub.loc[idx]
        return {
            "구분": row[col_group],
            "폐기물 종류": row[col_name],
            "처리 방법": row[col_method],
            "score": sc
        }

    # 1순위: '구분'에 '처리기준' 포함
    mask = df_ref[col_group].str.contains("처리기준", case=False, na=False)
    best = pick_best(df_ref[mask])
    if best:
        return best
    # 2순위: 전체 중 1건
    return pick_best(df_ref)

# =========================================
# Streamlit App
# =========================================
st.set_page_config(page_title=APP_TITLE, page_icon="♻️", layout="wide")

# 세션 상태 초기화
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False
    st.session_state.df_main = None
    st.session_state.df_ref = None
    st.session_state.xlsx_path = None

if "OPENAI_API_KEY" not in st.session_state:
    try:
        secret_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        secret_key = None
    env_key = os.getenv("OPENAI_API_KEY")
    st.session_state.OPENAI_API_KEY = secret_key or env_key

def load_app_data(xlsx_path: str):
    try:
        df_main, df_ref = load_data(xlsx_path)
        st.session_state.df_main = df_main
        st.session_state.df_ref = df_ref
        st.session_state.xlsx_path = xlsx_path
        st.session_state.data_loaded = True
    except Exception as e:
        st.error(f"데이터 로딩 오류: {e}")
        with st.expander("상세 오류"):
            st.code(traceback.format_exc())

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("🔧 시스템 관리")

    if st.button("📊 데이터 불러오기/새로고침", use_container_width=True):
        st.session_state.data_loaded = False
        load_app_data(DEFAULT_XLSX)
        st.rerun()

    if st.session_state.data_loaded:
        st.success("✅ 데이터 로드됨")
        st.info(f"📂 경로: {st.session_state.xlsx_path}")
    else:
        st.warning("⚠️ 데이터 미로드")

    st.subheader("🤖 OpenAI 연결 설정")
    if st.session_state.get("OPENAI_API_KEY"):
        st.success("🔑 OpenAI 키가 설정되어 있습니다 (Secrets 또는 Env).")
        st.caption("사이드바 입력은 Secrets/Env가 없을 때만 사용됩니다.")
        api_key_input = ""
    else:
        api_key_input = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        if api_key_input:
            st.session_state["OPENAI_API_KEY"] = api_key_input
            st.success("🔑 OpenAI 키가 세션에 설정되었습니다.")

# ---------------- Main ----------------
st.title(APP_TITLE)
st.markdown("---")

if not st.session_state.data_loaded:
    with st.spinner("📊 데이터를 로딩중입니다..."):
        load_app_data(DEFAULT_XLSX)

if st.session_state.data_loaded:
    st.subheader("🔍 폐기물 정보 입력")
    col1, col2, col3 = st.columns([2, 1, 2])
    with col1:
        waste_name = st.text_input("폐기물명 *", placeholder="예: 폐유, 폐페인트 슬러지")
    with col2:
        phase = st.selectbox("성상", ["선택안함", "고체", "액체"])
    with col3:
        material = st.text_input("재질", placeholder="예: PET, 고무, 유기용제 함유")

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

                    # OpenAI client
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
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.info(f"**매칭된 폐기물**: {best_row['폐기물 종류']}")
                        with c2:
                            st.info(f"**매칭 방식**: {debug_info.get('match_type')} (유사도: {debug_info.get('score', 0):.1f}%)")
                        with c3:
                            dept = get_dept_from_row(best_row)
                            st.info(f"**부서**: {dept or '정보없음'}")

                        # 부서가 '환경자원그룹'일 때만 법령참고 1건 노출
                        if dept == "환경자원그룹":
                            ref = find_best_ref_one(st.session_state.df_ref, best_row["폐기물 종류"])
                            if ref:
                                st.markdown("---")
                                st.subheader("📖 법령 참고 (시행규칙 별표5 · 처리기준 및 방법 · 상위 1건)")
                                st.markdown(f"- **구분**: {ref['구분']}")
                                st.markdown(f"- **폐기물 종류**: {ref['폐기물 종류']}")
                                st.markdown(f"- **처리 기준·방법**: {ref['처리 방법']}")
                            else:
                                st.caption("법령참고: 일치 항목 없음")

                    else:
                        # 매칭 실패 → OpenAI로 제안
                        st.error("❌ 일치하는 항목을 찾지 못했습니다.")
                        if debug_info.get("candidates"):
                            st.subheader("💡 유사 후보(Top)")
                            for i, c in enumerate(debug_info["candidates"][:5], 1):
                                st.markdown(f"**{i}. {c['폐기물 종류']}** (유사도: {c['score']:.1f}%)")
                                st.markdown(f" 처리 방법: {c['처리 방법']}")

                        if openai_client:
                            try:
                                sys_msg = (
                                    "너는 한국의 폐기물관리법 및 하위법령을 잘 아는 전문가다. "
                                    "입력된 폐기물명/성상/재질을 바탕으로 ‘가능성 높은 처리방법’과 "
                                    "검토 포인트를 한국어로 간결히 제안하라. "
                                    "사내 기준이 아님을 명시하지 말고, 법령 일반 원칙 수준에서만 답하라. "
                                    "목록 3개 이내로."
                                )
                                user_msg = {
                                    "폐기물명": waste_name,
                                    "성상": phase_input or "",
                                    "재질": material
                                }
                                rsp = openai_client.responses.create(
                                    model="gpt-4o-mini",
                                    input=[
                                        {"role": "system", "content": sys_msg},
                                        {"role": "user", "content": str(user_msg)}
                                    ],
                                    temperature=0.2
                                )
                                suggestion = rsp.output_text.strip()
                                st.markdown("### 🤔 OpenAI 제안(참고용)")
                                st.warning(
                                    "이 제안은 모델 생성 결과로 **환각(사실과 다른 내용) 가능성**이 있습니다. "
                                    "반드시 사내 기준 및 법적 근거를 재검토하세요."
                                )
                                st.markdown(suggestion if suggestion else "- 제안 생성 실패")
                            except Exception as e:
                                st.caption(f"OpenAI 제안 실패: {e}")

                except Exception as e:
                    st.error(f"❌ 검색 중 오류가 발생했습니다: {str(e)}")
                    with st.expander("상세 오류 정보"):
                        st.code(traceback.format_exc())
else:
    st.error("❌ 데이터를 로드할 수 없습니다. '데이터 불러오기/새로고침'을 클릭하거나 파일을 확인하세요.")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>"
    "💡 시스템 문의: 환경자원그룹 자원재활용섹션 (790-8526) | "
    "📧 실제 업무 적용은 사내 기준을 우선하세요"
    "</div>",
    unsafe_allow_html=True
)
