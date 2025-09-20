# app.py
import os
import re
import unicodedata
import traceback
from typing import Optional, List, Tuple, Dict, Any

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz

APP_TITLE = "사내 폐기물 처리방법 조회"
DEFAULT_XLSX = "wasteinfo.xlsx"

# ---------- 유틸 ----------
def normalize_korean(s: str) -> str:
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def canon(s: str) -> str:
    """비교용 정규화: 한글 NFKC, 소문자, 공백/하이픈 제거, 괄호 통일"""
    t = normalize_korean(s).lower()
    t = t.replace(" ", "")
    t = t.replace("-", "")
    t = t.replace("[", "(").replace("]", ")")
    return t

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
        if req not in df_main.columns:
            raise ValueError(f"메인 시트 '{req}' 컬럼 없음")
    for req in ["구분", "폐기물 종류", "처리 방법"]:
        if req not in df_ref.columns:
            raise ValueError(f"참고 시트 '{req}' 컬럼 없음")

    df_main["폐기물 종류"] = df_main["폐기물 종류"].astype(str).str.strip()
    df_main["처리 방법"]   = df_main["처리 방법"].astype(str).str.strip()
    df_ref["구분"]        = df_ref["구분"].astype(str).str.strip()
    df_ref["폐기물 종류"]  = df_ref["폐기물 종류"].astype(str).str.strip()
    df_ref["처리 방법"]    = df_ref["처리 방법"].astype(str).str.strip()
    return df_main, df_ref

# ---------- 매칭 ----------
def normalize_query(name: str, phase: Optional[str], material: str, openai_client=None) -> List[str]:
    base_terms = []
    for s in [name, material, phase or ""]:
        s = (s or "").strip()
        if s:
            base_terms.append(normalize_korean(s))
    candidates = list(dict.fromkeys([t for t in base_terms if t]))
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

    for q in query_terms:
        exact = df_main[df_main[col_name].str.strip().str.lower() == q.strip().lower()]
        if len(exact) > 0:
            row = exact.iloc[0]
            debug["match_type"] = "정확 일치"; debug["score"] = 100.0
            return row, debug
        part = df_main[df_main[col_name].str.contains(q, case=False, na=False)]
        if len(part) > 0:
            row = part.iloc[0]
            debug["match_type"] = "부분 일치"; debug["score"] = 95.0
            return row, debug

    all_scores: List[Tuple[int, float]] = []
    for q in query_terms:
        all_scores.extend(_score_series(df_main[col_name], q))
    all_scores.sort(key=lambda x: -x[1])
    top = all_scores[:10]

    best_row, best_score = None, 0.0
    seen = set()
    for idx, sc in top:
        if idx in seen: continue
        seen.add(idx)
        row = df_main.loc[idx]
        debug["candidates"].append({"폐기물 종류": row[col_name], "처리 방법": row[col_method], "score": sc})
        if sc > best_score:
            best_row, best_score = row, sc
    if best_row is not None and best_score >= threshold:
        debug["match_type"] = "퍼지 매칭"; debug["score"] = best_score
        return best_row, debug
    return None, debug

def find_refs(df_ref: pd.DataFrame, keyword: str, topk: int = 1) -> list[dict]:
    col_name, col_method, col_group = "폐기물 종류", "처리 방법", "구분"
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
        out.append({"구분": row[col_group], "폐기물 종류": row[col_name], "처리 방법": row[col_method], "score": sc})
    return out

# ---------- Streamlit ----------
st.set_page_config(page_title=APP_TITLE, page_icon="♻️", layout="wide")

if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False
    st.session_state.df_main = None
    st.session_state.df_ref = None
    st.session_state.xlsx_path = None

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
        load_app_data(DEFAULT_XLSX)
        st.rerun()
    if st.session_state.data_loaded:
        st.success("✅ 데이터 로드됨"); st.info(f"📂 경로: {st.session_state.xlsx_path}")
    else:
        st.warning("⚠️ 데이터 미로드")

st.title(APP_TITLE); st.markdown("---")

if not st.session_state.data_loaded:
    with st.spinner("📊 데이터를 로딩중입니다..."):
        load_app_data(DEFAULT_XLSX)

if st.session_state.data_loaded:
    dfm = st.session_state.df_main
    COL_WASTE = "폐기물 종류"
    COL_METHOD = "처리 방법"

    # 허용/배제 키
    ALLOW_KEY = "환경자원그룹(790-8526)"
    EXCLUDE_KEYS = [
        "중앙야적장 (광양자재지원 790-2732)",
        "제강부 스크랩장 (삼진기업 790-2815)",
    ]

    st.subheader("🔍 폐기물 정보 입력")
    c1, c2, c3 = st.columns([2, 1, 2])
    with c1:
        waste_name = st.text_input("폐기물명 *", placeholder="예: 폐유, 폐페인트 슬러지")
    with c2:
        phase = st.selectbox("성상", ["선택안함", "고체", "액체"])
    with c3:
        material = st.text_input("재질", placeholder="예: PET, 고무, 유기용제 함유")

    st.markdown("---")
    center = st.columns([1, 2, 1])[1]
    with center:
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

                    if best_row is not None:
                        st.subheader("✅ 처리 방법 (사내 기준)")
                        st.success(f"**처리 방법**: {best_row[COL_METHOD]}")
                        cA, cB = st.columns(2)
                        with cA:
                            st.info(f"**매칭된 폐기물**: {best_row[COL_WASTE]}")
                        with cB:
                            st.info(f"**매칭 방식**: {debug_info.get('match_type')} (유사도: {debug_info.get('score', 0):.1f}%)")

                        # ---------- 법령 표출 조건: 처리방법만 사용 ----------
                        method_val = str(best_row[COL_METHOD])
                        show_law = contains_key(method_val, ALLOW_KEY) and not any(
                            contains_key(method_val, k) for k in EXCLUDE_KEYS
                        )

                        if show_law:
                            refs = find_refs(st.session_state.df_ref, best_row[COL_WASTE], topk=1)
                            if refs:
                                st.markdown("---")
                                st.subheader("📖 법령 참고 (시행규칙 별표5, 1건)")
                                ref = refs[0]
                                st.markdown(f"- **구분**: {ref['구분']}")
                                st.markdown(f"- **폐기물 종류**: {ref['폐기물 종류']}")
                                st.markdown(f"- **처리 방법**: {ref['처리 방법']}")
                            else:
                                st.info("별표5에서 연관 항목을 찾지 못했습니다.")
                        else:
                            st.caption("법령 참고 숨김: 조건 미충족(환경자원그룹 아니거나 중앙야적장/스크랩장)")

                    else:
                        st.error("❌ 일치 항목을 찾지 못했습니다.")
                        if debug_info.get("candidates"):
                            st.subheader("💡 유사한 항목들(Top)")
                            for i, c in enumerate(debug_info["candidates"][:5], 1):
                                st.markdown(f"**{i}. {c['폐기물 종류']}** (유사도: {c['score']:.1f}%)")
                                st.markdown(f"   처리 방법: {c['처리 방법']}")
                except Exception as e:
                    st.error(f"❌ 검색 중 오류: {str(e)}")
                    with st.expander("상세 오류 정보"):
                        st.code(traceback.format_exc())
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
