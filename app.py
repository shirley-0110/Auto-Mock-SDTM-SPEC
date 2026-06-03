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




# =========================================================================================================================================================
# 基本工具函式
# =========================================================================================================================================================
def normalize_text(x):
    if pd.isna(x):
        return ""
    x = str(x) #統一資料型態
    x = x.replace("\n", " ") #移除換行
    x = x.replace("\r", " ")
    x = x.replace("\xa0", " ")
    x = re.sub(r"\s+", " ", x) #壓縮多於空白
    return x.strip().upper()
    # End=========================================================

def normalize_columns(df):
    cols = []
    for c in df.columns:
        c = str(c)
        c = c.replace("\n", " ")
        c = c.replace("\r", " ")
        c = c.replace("\xa0", " ")
        c = re.sub(r"\s+", " ", c).strip()
        cols.append(c)
    df.columns = cols
    return df
    # End=========================================================


def find_column(columns, required_keywords):
    for col in columns:
        upper_col = normalize_text(col)
        if all(k.upper() in upper_col for k in required_keywords):
            return col
    return None
    # End=========================================================


def row_contains_keywords(row_values, keyword_groups):
    cells = [normalize_text(v) for v in row_values]

    for cell in cells:
        for group in keyword_groups:
            if all(k.upper() in cell for k in group):
                return True
    return False
    # End=========================================================


def detect_header_row(file_bytes, sheet_name, keyword_groups, max_scan_rows=30):
    raw_df = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        nrows=max_scan_rows,
        dtype=str
    )

    for idx, row in raw_df.iterrows():
        if row_contains_keywords(row.tolist(), keyword_groups):
            return idx

    return None
    # End=========================================================


def read_sheet_with_detected_header(
    file_bytes,
    sheet_name,
    keyword_groups,
    manual_header_row_excel=None,
    max_scan_rows=30
):
    if manual_header_row_excel is not None:
        header_row_zero_based = manual_header_row_excel - 1
    else:
        header_row_zero_based = detect_header_row(
            file_bytes=file_bytes,
            sheet_name=sheet_name,
            keyword_groups=keyword_groups,
            max_scan_rows=max_scan_rows
        )

    if header_row_zero_based is None:
        raise ValueError(f"無法自動判斷 {sheet_name} 的 header row")

    df = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=header_row_zero_based
    )
    df = normalize_columns(df)

    return df, header_row_zero_based + 1
    # End=========================================================




def build_soa_visit_list(
    file_bytes,
    manual_soa_header=None,
    manual_folder_header=None
):
    """
    從 SoA + Folder 建立 SoA List:
      CRF Dataset / Abbreviation / Visit

    規則：
      - SoA 的 row = Source CRF Sheet (Form OID)
      - SoA 的 visit 欄位只要 cell = X，就輸出一列
      - Folder 的 Abbreviation -> Full Term 對出 Visit
    """
    # -----------------------------
    # 1) 讀 SoA
    # -----------------------------
    soa_df, _ = read_sheet_with_detected_header(
        file_bytes=file_bytes,
        sheet_name="SoA",
        keyword_groups=[["FORM", "OID"]],
        manual_header_row_excel=manual_soa_header
    )

    form_oid_col = find_column(soa_df.columns, ["FORM", "OID"])
    if form_oid_col is None:
        raise ValueError("SoA 分頁中找不到 Form OID 欄位")

    # SoA 所有欄位
    soa_columns = [str(c).strip() for c in soa_df.columns if str(c).strip()]

    # 這些欄位不是 visit abbreviation
    non_visit_headers = {
        "FORM OID",
        "FORMOID",
        "FORM",
        "CRF NAME",
        "FORM NAME",
        "DESCRIPTION",
        "SEQ",
        "ORDER"
    }

    visit_cols = []
    for idx, col in enumerate(soa_columns):
        col_up = normalize_text(col)
        
        if col_up not in non_visit_headers:
            visit_cols.append((col, idx))

    # -----------------------------
    # 2) 讀 Folder
    # -----------------------------
    folder_df, _ = read_sheet_with_detected_header(
        file_bytes=file_bytes,
        sheet_name="Folder",
        keyword_groups=[["ABBREVIATION"], ["FULL", "TERM"]],
        manual_header_row_excel=manual_folder_header
    )

    abbr_col = find_column(folder_df.columns, ["ABBREVIATION"])
    if abbr_col is None:
        raise ValueError("Folder 分頁中找不到 Abbreviation 欄位")

    full_term_col = find_column(folder_df.columns, ["FULL", "TERM"])
    if full_term_col is None:
        raise ValueError("Folder 分頁中找不到 Full Term 欄位")

    folder_work = folder_df[[abbr_col, full_term_col]].copy()
    folder_work.columns = ["Abbreviation", "Visit"]

    folder_work["Abbreviation"] = (
        folder_work["Abbreviation"].fillna("").astype(str).str.strip().str.upper()
    )
    folder_work["Visit"] = (
        folder_work["Visit"].fillna("").astype(str).str.strip()
    )

    folder_work = folder_work[
        (folder_work["Abbreviation"] != "") &
        (folder_work["Visit"] != "")
    ].drop_duplicates(subset=["Abbreviation"], keep="first")

    folder_lookup = dict(zip(folder_work["Abbreviation"], folder_work["Visit"]))

    # -----------------------------
    # 3) 展開 SoA List
    # -----------------------------
    records = []

    for _, row in soa_df.iterrows():
        source_sheet = str(row.get(form_oid_col, "")).strip().upper()

        if not source_sheet:
            continue

        for col, col_idx in visit_cols:
            abbr = normalize_text(col)
            abbr = re.sub(r'[\*\^]+', '', abbr).strip()
            cell_val = str(row.get(col, "")).strip().upper()

            # 只抓 ticked X
            if cell_val == "X":
                visit_name = folder_lookup.get(abbr, "")

                records.append({
                    "CRF Dataset": source_sheet,
                    "Abbreviation": abbr,
                    "Visit": visit_name,
                    "Visit_order": col_idx
                })

    if records:
        out_df = pd.DataFrame(records).drop_duplicates()
        
        # 排序：先 CRF Dataset，再 Visit_order
        out_df = out_df.sort_values(
            by=["CRF Dataset", "Visit_order"],
            ascending=[True, True]
        ).reset_index(drop=True)

    else:
        out_df = pd.DataFrame(columns=[
            "CRF Dataset", "Abbreviation", "Visit", "Visit_order"
        ])

    return out_df
    # End=========================================================




def build_tv_from_soa_list(
    soa_list_df,
    protocol_no="",
    ordered_columns=None
):
    """
    從 SoA List 建立 TV domain

    規則：
      - 依 Abbreviation 去重
      - 保留 SoA 原順序
      - 排除 Source CRF Sheet == Abbreviation
    """

    if ordered_columns is None:
        ordered_columns = [
            "STUDYID", "DOMAIN", "VISITNUM", "VISIT", "VISITDY",
            "ARMCD", "ARM", "TVSTRL", "TVENRL"
        ]

    def make_row():
        row = {c: "" for c in ordered_columns}
        if "STUDYID" in row:
            row["STUDYID"] = protocol_no
        if "DOMAIN" in row:
            row["DOMAIN"] = "TV"
        return row

    if soa_list_df is None or soa_list_df.empty:
        return pd.DataFrame([make_row()], columns=ordered_columns)

    df = soa_list_df.copy()
    
    if "visit_order" not in df.columns:
        df["visit_order"] = range(len(df))

    for c in ["Source CRF Sheet", "Abbreviation", "Visit"]:
        if c not in df.columns:
            df[c] = ""

    df["Source CRF Sheet"] = df["Source CRF Sheet"].astype(str).str.upper().str.strip()
    df["Abbreviation"] = df["Abbreviation"].astype(str).str.upper().str.strip()
    df["Visit"] = df["Visit"].astype(str).str.strip()

    # 1 移除 Source CRF Sheet == Abbreviation
    df = df[df["Source CRF Sheet"] != df["Abbreviation"]]

    # 2 只保留有 Visit（Folder 已對應）
    df = df[df["Visit"] != ""]

    # 3 先排序（依 SoA 欄位順序）、去重
    df["visit_order"] = pd.to_numeric(df["visit_order"], errors="coerce")
    df = df.sort_values(by="visit_order")
    df = df.drop_duplicates(subset=["Abbreviation"], keep="first")

    # 4 建立 TV rows
    rows = []

    for _, r in df.iterrows():
        row = make_row()

        if "VISIT" in row:
            row["VISIT"] = r["Visit"]

        rows.append(row)

    if not rows:
        rows = [make_row()]

    return pd.DataFrame(rows, columns=ordered_columns)







# =========================================================================================================================================================
# 主流程 UI
# =========================================================================================================================================================

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


        # 呼叫
        soa_list_df = build_soa_visit_list(
            file_bytes=file_bytes,
            manual_soa_header=None,
            manual_folder_header=None
        )
        
        unique_visit_df = (
            soa_df[["Abbreviation", "Visit", "Visit_order"]]
            .drop_duplicates()
            .sort_values("Visit_order")
            .reset_index(drop=True)
        )
        st.write(unique_visit_df)
        
        
        # -------------------------------------------------
        # Step 1：CRF → SDTM Mapping
        # -------------------------------------------------
        st.markdown("## Step 1｜CRF → SDTM Mapping")


    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
