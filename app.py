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
# GitHub repo 루트에 wasteinfo.xlsx 업로드 필요
DEFAULT_XLSX = "wasteinfo.xlsx"

# =========================================
# 유틸
# =========================================
def normalize_korean(s: str) -> str:
    """NFKC 정규화 + 공백 정리"""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

# =========================================
# 데이터 로딩
# =========================================
def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]  # '처리 방법 ' → '처리 방법'
    return df

@st.cache_data(show_spinner=False)
def load_data(xlsx_path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(xlsx_path)

    df_main = pd.read_excel(xlsx_path, sheet_name="광양소 폐기물처리방법", engine="openpyxl")
    df_ref  = pd.read_excel(xlsx_path, sheet_name="폐기물관리법_시행규칙_별표5", engine="openpyxl")

    df_main = _normalize_columns(df_main)
    df_ref  = _normalize_columns(df_ref)

    return df_main, df_ref

# =========================================
# 매칭/검색
# =========================================
def normalize_query(name: str, phase: Optional[str], material: str) -> List[str]:
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

def search_best(df_main: pd.DataFrame, query_terms: List[str], threshold: int = 60):
    col_name, col_method = "폐기물 종류", "처리 방법"
    debug: Dict[str, Any] = {"match_type": "", "score": 0.0, "candidates": []}

    # 정확/부분 일치
    for q in query_terms:
        exact = df_main[df_main[col_name].str.strip().str.lower() == q.strip().lower()]
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

    # 퍼지 매칭
    all_scores = []
    for q in query_terms:
        all_scores.extend(_score_series(df_main[col_name], q))
    all_scores.sort(key=lambda x: -x[1])
    top = all_scores[:10]

    debug["candidates"] = []
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

def find_refs(df_ref: pd.DataFrame, keyword: str, topk: int = 3):
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
        st.success("✅ 데이터 로드됨")
        st.info(f"📂 경로: {st.session_state.xlsx_path}")
    else:
        st.warning("⚠️ 데이터 미로드")

st.title(APP_TITLE)
st.markdown("---")

if not st.session_state.data_loaded:
    with st.spinner("📊 데이터를 로딩중입니다..."):
        load_app_data(DEFAULT_XLSX)

if st.session_state.data_loaded:
    waste_name = st.text_input("폐기물명 *", placeholder="예: 폐유, 폐페인트 슬러지")
    phase = st.selectbox("성상", ["선택안함", "고체", "액체"])
    material = st.text_input("재질", placeholder="예: PET, 고무, 유기용제 함유")

    if st.button("🔍 처리방법 조회", type="primary"):
        if not waste_name.strip():
            st.warning("⚠️ 폐기물명을 입력해주세요.")
        else:
            query_terms = normalize_query(waste_name, None if phase == "선택안함" else phase, material)
            best_row, debug_info = search_best(st.session_state.df_main, query_terms)

            if best_row is not None:
                st.subheader("✅ 처리 방법 (사내 기준)")
                st.success(f"**처리 방법**: {best_row['처리 방법']}")
                st.info(f"**매칭된 폐기물**: {best_row['폐기물 종류']} / {debug_info.get('match_type')}")

                refs = find_refs(st.session_state.df_ref, best_row["폐기물 종류"])
                if refs:
                    st.markdown("---")
                    st.subheader("📖 법령 참고 (폐기물관리법 시행규칙 별표5)")
                    for i, ref in enumerate(refs[:3], 1):
                        st.markdown(f"**{i}. {ref['구분']}**")
                        st.markdown(f"- 폐기물 종류: {ref['폐기물 종류']}")
                        st.markdown(f"- 처리 방법: {ref['처리 방법']}")
            else:
                st.error("❌ 일치하는 항목을 찾지 못했습니다.")
