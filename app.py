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
DEFAULT_XLSX = "wasteinfo.xlsx"   # GitHub repo 루트에 wasteinfo.xlsx 업로드 필요
FUZZ_THRESHOLD = 60               # 유사도 하한(%) — 미만이면 OpenAI 제안 경로로 전환

# =========================================
# 유틸
# =========================================
def normalize_korean(s: str) -> str:
    """NFKC 정규화 + 공백 정리"""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def _get_dept_value(row: pd.Series) -> str:
    """행에서 '부서'에 해당하는 값을 유연하게 탐색"""
    dept_cols = ["부서", "담당 부서", "처리부서", "부서명"]
    for c in dept_cols:
        if c in row.index:
            return str(row[c] or "").strip()
    return ""

def _is_env_group(row: pd.Series) -> bool:
    """부서가 '환경자원그룹'인 경우만 True"""
    dept = _get_dept_value(row)
    return "환경자원그룹" in dept

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

    # 부서 컬럼 존재 시 문자열화
    for c in ["부서", "담당 부서", "처리부서", "부서명"]:
        if c in df_main.columns:
            df_main[c] = df_main[c].astype(str).str.strip()

    return df_main, df_ref

# =========================================
# 매칭/검색
# =========================================
def normalize_query(name: str, phase: Optional[str], material: str, openai_client=None) -> List[str]:
    """
    기본 후보: [폐기물명, 재질, 성상]
    OpenAI가 있으면 동의어/오타 보정 후보를 JSON 배열로 받아 추가.
    """
    base_terms = []
    for s in [name, material, phase or ""]:
        s = (s or "").strip()
        if s:
            base_terms.append(normalize_korean(s))

    candidates = list(dict.fromkeys([t for t in base]()
