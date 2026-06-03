import streamlit as st
import pandas as pd
import re
import os
import hashlib
import io

from io import BytesIO

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from difflib import get_close_matches

# Step 2 用到 sas7bdat
try:
    import pyreadstat
    HAS_PYREADSTAT = True
except Exception:
    HAS_PYREADSTAT = False



# =========================================================
# 主流程 UI
# =========================================================
st.set_page_config(page_title="Auto SDTM SPEC", layout="wide")
st.title("Auto SDTM SPEC")



uploaded_file = st.file_uploader("請上傳 CRF Mapping Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()

    current_upload_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("current_upload_key") != current_upload_key:
        st.session_state["current_upload_key"] = current_upload_key
        st.session_state["run_step2"] = False
        st.session_state["step1_cache_key"] = None
        st.session_state["step1_result"] = None

    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
        all_sheets = xls.sheet_names

        # -------------------------------------------------
        # Header Override：放在上傳檔案下面
        # -------------------------------------------------
        st.markdown("### Header Override（選填）")

        col1, col2 = st.columns(2)

        with col1:
            use_manual_soa_header = st.checkbox("手動指定 SoA header row")
            manual_soa_header = None
            if use_manual_soa_header:
                manual_soa_header = st.number_input(
                    "SoA header 在 Excel 第幾列？",
                    min_value=1,
                    value=2,
                    step=1
                )

        with col2:
            use_manual_domain_header = st.checkbox("所有 Domain Sheet 使用同一個 header row")
            common_domain_header = None
            if use_manual_domain_header:
                common_domain_header = st.number_input(
                    "所有 Domain Sheet header 在 Excel 第幾列？",
                    min_value=1,
                    value=2,
                    step=1
                )

        # -------------------------------------------------
        # Step 1：CRF → SDTM Mapping
        # -------------------------------------------------
        st.markdown("## Step 1｜CRF → SDTM Mapping")


    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
