# app.py
import os, re, unicodedata, traceback
from typing import Optional, List, Tuple, Dict, Any

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

APP_TITLE = "사내 폐기물 처리방법 조회"
DEFAULT_XLSX = "wasteinfo.xlsx"

# ---------- 유틸 ----------
def normalize_korean(s: str) -> str:
    s = unicodedata.normalize("NFKC", str(s or ""))
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def canon(s: str) -> str:
    t = normalize_korean(s).lower()
    return t.replace(" ", "").replace("-", "").replace("[", "(").replace("]", ")")

def contains_key(text: str, key: str) -> bool:
    return canon(key) in canon(text or "")

# ---------- 데이터 ----------
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
    for req in ["폐기물 종류", "처리 방법"]:
        if req not in df_main.columns: raise ValueError(f"메인 시트 '{req}' 컬럼 없음")
    for req in ["구분", "폐기물 종류", "처리 방법"]:
        if req not in df_ref.columns: raise ValueError(f"참고 시트 '{req}' 컬럼 없음")
    df_main["폐기물 종류"] = df_main["폐기물 종류"].astype(str).str.strip()
    df_main["처리 방법"]   = df_main["처리 방법"].astype(str).str.strip()
    df_ref["구분"]        = df_ref["구분"].astype(str).str.strip()
    df_ref["폐기물 종류"]  = df_ref["폐기물 종류"].astype(str).str.strip()
    df_ref["처리 방법"]    = df_ref["처리 방법"].astype(str).str.strip()
    return df_main, df_ref

# ---------- 매칭 ----------
def normalize_query(name: str, phase: Optional[str], material: str, openai_client=None) -> List[str]:
    base_terms = [normalize_korean(x) for x in [name, material, phase or ""] if str(x).strip()]
    return list(dict.fromkeys(base_terms)) or [normalize_korean(name)]

def _score_series(series: pd.Series, query: str) -> List[Tuple[int, float]]:
    qn = normalize_korean(query); scores = []
    for idx, val in series.items():
        score = fuzz.WRatio(qn, normalize_korean(val))
        scores.append((idx, float(score)))
    return scores

def search_best(df_main: pd.DataFrame, query_terms: List[str], threshold: int = 60) -> Tuple[Optional[pd.Series], Dict[str, Any]]:
    col_name, col_method = "폐기물 종류", "처리 방법"
    debug: Dict[str, Any] = {"match_type": "", "score": 0.0, "candidates": []}
    # 정확/부분
    for q in query_terms:
        exact = df_main[df_main[col_name].str.strip().str.lower() == q.strip().lower()]
        if len(exact) > 0:
            row = exact.iloc[0]; debug["match_type"] = "정확 일치"; debug["score"] = 100.0
            return row, debug
        part = df_main[df_main[col_name].str.contains(q, case=False, na=False)]
        if len(part) > 0:
            row = part.iloc[0]; debug["match_type"] = "부분 일치"; debug["score"] = 95.0
            return row, debug
    # 퍼지
    all_scores: List[Tuple[int, float]] = []
    for q in query_terms: all_scores.extend(_score_series(df_main[col_name], q))
    all_scores.sort(key=lambda x: -x[1]); top = all_scores[:10]
    best_row, best_score = None, 0.0; seen = set()
    for idx, sc in top:
        if idx in seen: continue
        seen.add(idx); row = df_main.loc[idx]
        debug["candidates"].append({"폐기물 종류": row[col_name], "처리 방법": row[col_method], "score": sc})
        if sc > best_score: best_row, best_score = row, sc
    if best_row is not None and best_score >= threshold:
        debug["match_type"] = "퍼지 매칭"; debug["score"] = best_score
        return best_row, debug
    return None, debug

# ---------- 법령: 정확 일치만 ----------
def get_ref_exact(df_ref: pd.DataFrame, waste_name: str) -> Optional[Dict[str, str]]:
    key = canon(waste_name)
    for _, r in df_ref.iterrows():
        if canon(r["폐기물 종류"]) == key:
            return {"구분": r["구분"], "폐기물 종류": r["폐기물 종류"], "처리 방법": r["처리 방법"]}
    return None

# ---------- OpenAI 제안 ----------
def build_openai():
    # Secrets → Env → 입력 순
    try: sec = st.secrets.get("OPENAI_API_KEY", None)
    except Exception: sec = None
    env = os.getenv("OPENAI_API_KEY")
    key = sec or env or st.session_state.get("OPENAI_API_KEY")
    if not key: return None, "OpenAI 키가 설정되지 않았습니다."
    try:
        from openai import OpenAI
        return OpenAI(api_key=key), None
    except Exception as e:
        return None, f"OpenAI 초기화 실패: {e}"

def propose_with_openai(client, waste_name: str, phase: Optional[str], material: str) -> str:
    sys = (
        "너는 대한민국 폐기물관리법 기반 조언가다. 모르면 모른다고 말한다. "
        "‘비공식 제안’ 경고를 반드시 포함한다. 출력은 5줄 이내 한국어 마크다운."
    )
    user = {
        "요청": "별표5에 정확 일치가 없어 참고용 처리방법 제안",
        "폐기물명": waste_name,
        "성상": phase or "",
        "재질/특성": material or "",
        "근거": "폐기물관리법 및 시행규칙, 별표5 범위 내에서 일반 원칙 중심(소각, 매립, 위탁처리 등)"
    }
    rsp = client.responses.create(
        model="gpt-4o-mini",
        temperature=0.2,
        input=[
            {"role": "system", "content": sys},
            {"role": "user", "content": str(user)}
        ],
    )
    return rsp.output_text.strip()

# ---------- Streamlit ----------
st.set_page_config(page_title=APP_TITLE, page_icon="♻️", layout="wide")

if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False
    st.session_state.df_main = None
    st.session_state.df_ref = None
    st.session_state.xlsx_path = None
if "OPENAI_API_KEY" not in st.session_state:
    st.session_state.OPENAI_API_KEY = ""

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

with st.sidebar:
    st.header("🔧 시스템 관리")
    if st.button("📊 데이터 불러오기/새로고침", use_container_width=True):
        st.session_state.data_loaded = False
        load_app_data(DEFAULT_XLSX); st.rerun()
    if st.session_state.data_loaded:
        st.success("✅ 데이터 로드됨"); st.info(f"📂 경로: {st.session_state.xlsx_path}")
    else:
        st.warning("⚠️ 데이터 미로드")
    st.subheader("🤖 OpenAI 연결(선택)")
    if not (os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY", None)):
        st.session_state.OPENAI_API_KEY = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")

st.title(APP_TITLE); st.markdown("---")

if not st.session_state.data_loaded:
    with st.spinner("📊 데이터를 로딩중입니다..."):
        load_app_data(DEFAULT_XLSX)

if st.session_state.data_loaded:
    dfm = st.session_state.df_main
    dfref = st.session_state.df_ref
    COL_WASTE, COL_METHOD = "폐기물 종류", "처리 방법"

    # 허용/배제 키
    ALLOW_KEY = "환경자원그룹(790-8526)"
    EXCLUDE_KEYS = [
        "중앙야적장 (광양자재지원 790-2732)",
        "제강부 스크랩장 (삼진기업 790-2815)",
    ]

    st.subheader("🔍 폐기물 정보 입력")
    c1, c2, c3 = st.columns([2, 1, 2])
    with c1: waste_name = st.text_input("폐기물명 *", placeholder="예: 폐유, 폐페인트 슬러지")
    with c2: phase = st.selectbox("성상", ["선택안함", "고체", "액체"])
    with c3: material = st.text_input("재질", placeholder="예: PET, 고무, 유기용제 함유")

    st.markdown("---")
    with st.columns([1,2,1])[1]:
        search_button = st.button("🔍 처리방법 조회", type="primary", use_container_width=True)

    if search_button:
        if not waste_name.strip():
            st.warning("⚠️ 폐기물명을 입력해주세요.")
        else:
            with st.spinner("🔍 검색중입니다..."):
                try:
                    phase_input = None if phase == "선택안함" else phase
                    query_terms = normalize_query(waste_name, phase_input, material, None)
                    best_row, debug_info = search_best(dfm, query_terms)

                    if best_row is None:
                        st.error("❌ 일치 항목을 찾지 못했습니다.")
                        if debug_info.get("candidates"):
                            st.subheader("💡 유사 항목(Top)")
                            for i, c in enumerate(debug_info["candidates"][:5], 1):
                                st.markdown(f"**{i}. {c['폐기물 종류']}** (유사도: {c['score']:.1f}%)")
                                st.markdown(f"   처리 방법: {c['처리 방법']}")
                    else:
                        st.subheader("✅ 처리 방법 (사내 기준)")
                        st.success(f"**처리 방법**: {best_row[COL_METHOD]}")
                        cA, cB = st.columns(2)
                        with cA: st.info(f"**매칭된 폐기물**: {best_row[COL_WASTE]}")
                        with cB: st.info(f"**매칭 방식**: {debug_info.get('match_type')} (유사도: {debug_info.get('score',0):.1f}%)")

                        # 표시 여부: 처리방법만 사용
                        method_val = str(best_row[COL_METHOD])
                        show_gate = contains_key(method_val, ALLOW_KEY) and not any(contains_key(method_val, k) for k in EXCLUDE_KEYS)

                        if show_gate:
                            # 1) 별표5 정확 일치만 표시
                            ref = get_ref_exact(dfref, best_row[COL_WASTE])
                            if ref:
                                st.markdown("---")
                                st.subheader("📖 법령 참고 (시행규칙 별표5, 정확 일치)")
                                st.markdown(f"- **구분**: {ref['구분']}")
                                st.markdown(f"- **폐기물 종류**: {ref['폐기물 종류']}")
                                st.markdown(f"- **처리 방법**: {ref['처리 방법']}")
                            else:
                                # 2) 정확 일치가 없으면 OpenAI 제안
                                client, err = build_openai()
                                st.markdown("---")
                                st.subheader("📝 참고 제안 (정확 일치 없음)")
                                if err:
                                    st.warning("OpenAI 미설정으로 제안 불가. 사이드바에서 API 키를 설정하세요.")
                                else:
                                    tip = propose_with_openai(client, waste_name, phase_input, material)
                                    st.markdown(tip)
                                    st.caption("⚠️ 비공식 참고용 제안입니다. 법적 효력 없음. 사내 기준과 법령 원문을 우선 검토하세요.")
                        else:
                            st.caption("법령 참고 숨김: 환경자원그룹이 아니거나 중앙야적장/스크랩장 관련입니다.")
                except Exception as e:
                    st.error(f"❌ 검색 중 오류: {str(e)}")
                    with st.expander("상세 오류 정보"): st.code(traceback.format_exc())
else:
    st.error("❌ 데이터를 로드할 수 없습니다. 사이드바에서 불러오기를 실행하세요.")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>"
    "💡 시스템 문의: 환경자원그룹 자원재활용섹션 (790-8526) | "
    "📧 실제 업무 적용은 사내 기준을 우선하세요"
    "</div>",
    unsafe_allow_html=True
)
