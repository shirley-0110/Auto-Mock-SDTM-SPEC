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



# ===================================================================================================================================================================================
# 所有 Function
# ===================================================================================================================================================================================

# =================================================================================================================
# 文字處理
# =================================================================================================================
def normalize_text(x):
    if pd.isna(x):
        return ""
    x = str(x) #統一資料型態
    x = x.replace("\n", " ").replace("\r", " ").replace("\xa0", " ") #移除換行
    x = re.sub(r"\s+", " ", x) #壓縮多於空白
    return x.strip().upper()
    # End=========================================================

def normalize_columns(df):
    df = df.copy()
    df.columns = [
        re.sub(r"\s+", " ", str(c).replace("\n", " ").replace("\r", " ").replace("\xa0", " ")).strip()
        for c in df.columns
    ]
    return df
    # End=========================================================


# =================================================================================================================
# 匯入Excel各種工具
# =================================================================================================================
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
    max_scan_rows=30
):
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




# 處理OID
def extract_form_oids(series):
    domains = set()

    for value in series.dropna():
        text = str(value).strip()
        if not text:
            continue

        parts = re.split(r"[,\n;/]+", text)

        for part in parts:
            item = part.strip()
            if item:
                domains.add(item.upper())

    return domains
    # End=========================================================



# 整個eCRF schema匯入+暫存
def build_step1_context(file_bytes, all_sheets):

    # 1. 匯入SoA
    soa_df, _ = read_sheet_with_detected_header(
        file_bytes=file_bytes,
        sheet_name="SoA",
        keyword_groups=[["FORM", "OID"]]
    )

    form_oid_col = find_column(soa_df.columns, ["FORM", "OID"])
    if form_oid_col is None:
        raise ValueError("SoA 找不到 Form OID")

    valid_domains = extract_form_oids(soa_df[form_oid_col])

    sheet_upper_map = {s.upper(): s for s in all_sheets}

    available_sheets = [
        sheet_upper_map[d] for d in valid_domains if d in sheet_upper_map
    ]

    missing_sheets = [
        d for d in valid_domains if d not in sheet_upper_map
    ]


    # 2. Folder
    folder_df, _ = read_sheet_with_detected_header(
        file_bytes=file_bytes,
        sheet_name="Folder",
        keyword_groups=[["ABBREVIATION"], ["FULL", "TERM"]]
    )


    # 3. CRF Domain Sheets
    domain_df_map = {}
    sheet_errors = []

    for sheet in available_sheets:

        try:
            df, _ = read_sheet_with_detected_header(
                file_bytes=file_bytes,
                sheet_name=sheet,
                keyword_groups=[["SDTM", "TARGET"]]
            )
            domain_df_map[sheet] = df

        except Exception:
            sheet_errors.append(sheet)

    return {
        "soa_df": soa_df,
        "folder_df": folder_df,
        "domain_df_map": domain_df_map,
        "available_sheets": available_sheets,
        "missing_sheets": missing_sheets,
        "sheet_errors": sheet_errors
    }
    # End=========================================================




# =================================================================================================================
# 特定使用
# =================================================================================================================

# 抓SoA的Visit
def build_soa_visit_list(soa_df, folder_df):
    """
    從 SoA + Folder 建立 SoA List:
      CRF Dataset / Abbreviation / Visit

    規則：
      - SoA 的 row = Source CRF Sheet (Form OID)
      - SoA 的 visit 欄位只要 cell = X，就輸出一列
      - Folder 的 Abbreviation -> Full Term 對出 Visit
    """

    # 1 呼叫SoA
    form_oid_col = find_column(soa_df.columns, ["FORM", "OID"])
    if form_oid_col is None:
        raise ValueError("SoA 分頁中找不到 Form OID 欄位")

    # SoA 所有欄位
    soa_columns = [str(c).strip() for c in soa_df.columns if str(c).strip()]

    # 這些欄位不是 visit abbreviation
    non_visit_headers = {
        "FORM OID", "FORMOID", "FORM",
        "CRF NAME", "FORM NAME",
        "DESCRIPTION", "SEQ", "ORDER"
    }

    visit_cols = [
        (col, idx)
        for idx, col in enumerate(soa_columns)
        if normalize_text(col) not in non_visit_headers
    ]
    
    # 2 呼叫 Folder
    abbr_col = find_column(folder_df.columns, ["ABBREVIATION"])
    if abbr_col is None:
        raise ValueError("Folder 分頁中找不到 Abbreviation 欄位")

    full_term_col = find_column(folder_df.columns, ["FULL", "TERM"])
    if full_term_col is None:
        raise ValueError("Folder 分頁中找不到 Full Term 欄位")

    folder_work = (
        folder_df[[abbr_col, full_term_col]]
        .rename(columns={abbr_col: "Abbreviation", full_term_col: "Visit"})
        .dropna()
    )
    
    folder_work["Abbreviation"] = folder_work["Abbreviation"].astype(str).str.strip().str.upper()
    folder_work["Visit"] = folder_work["Visit"].astype(str).str.strip()

    folder_lookup = dict(
        folder_work.drop_duplicates("Abbreviation")[["Abbreviation", "Visit"]].values
    )

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



# 抓各個CRF Domain的Field OID
def find_source_variable_column(columns):
    """
    Source CRF Variable 優先抓 Field OID
    """
    priority_exact = [
        "FIELD OID",
        "FIELDOID",
        "FIELD OID NAME",
        "CRF FIELD OID",
        "SOURCE FIELD OID",
        "VARIABLE",
        "VARIABLE NAME",
        "CRF VARIABLE",
        "SOURCE VARIABLE"
    ]

    normalized_map = {col: normalize_text(col) for col in columns}

    for target in priority_exact:
        for col, norm_col in normalized_map.items():
            if norm_col == target:
                return col

    for col, norm_col in normalized_map.items():
        if "FIELD" in norm_col and "OID" in norm_col:
            return col

    for col, norm_col in normalized_map.items():
        if "VARIABLE" in norm_col and "TARGET" not in norm_col and "SDTM" not in norm_col:
            return col

    return None
    # End=========================================================



# 處理SDTM IG Target
def parse_sdtm_targets(value):
    """
    規則：
      - 只用分號 ; 和換行切
      - 不用逗號和斜線切
      - 支援：
          AE.AETERM
          VS.VSTESTCD="TEMP"
          DM.SEX='F'
    """
    parsed_records = []
    unparsed_tokens = []

    if pd.isna(value):
        return parsed_records, unparsed_tokens

    text = str(value).strip()
    if not text:
        return parsed_records, unparsed_tokens

    tokens = re.split(r"[;\n]+", text)

    pattern = re.compile(
        r'^\s*([A-Za-z][A-Za-z0-9]{0,7})\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)\s*(?:=\s*["\']?(.*?)["\']?)?\s*$'
    )

    for token in tokens:
        token = token.strip()
        if not token:
            continue

        match = pattern.match(token)
        if match:
            dom, var, assign_val = match.groups()

            if assign_val is None:
                assign_val = ""
            else:
                assign_val = str(assign_val).strip()

            parsed_records.append({
                "SDTM Domain": dom.upper(),
                "SDTM Variable": var.upper(),
                "Assign Value": assign_val
            })
        else:
            unparsed_tokens.append(token)

    return parsed_records, unparsed_tokens
    # End=========================================================



# 處理CRF -> SDTM Variable Mapping
def build_sdtm_mapping(domain_df_map):

    mapping_records = []
    detail_records = []
    sheet_errors = []
    unparsed_records = []

    for sheet, df in domain_df_map.items():
        target_col = find_column(df.columns, ["SDTM", "TARGET"])
        source_var_col = find_source_variable_column(df.columns)

        if target_col is None:
            sheet_errors.append(sheet)
            continue

        for _, row in df.iterrows():
            raw_target = row[target_col]
            source_var = row[source_var_col] if source_var_col is not None else ""

            parsed_records, unparsed_tokens = parse_sdtm_targets(raw_target)

            for rec in parsed_records:
                mapping_records.append({
                    "SDTM Domain": rec["SDTM Domain"],
                    "SDTM Variable": rec["SDTM Variable"]
                })

                detail_records.append({
                    "CRF Dataset": sheet,
                    "CRF Variable": source_var,
                    "SDTM Domain": rec["SDTM Domain"],
                    "SDTM Variable": rec["SDTM Variable"],
                    "Assign Value": rec["Assign Value"],
                    "SDTM IG Target Raw": raw_target
                })

            for token in unparsed_tokens:
                if str(token).strip():
                    unparsed_records.append({
                        "CRF Dataset": sheet,
                        "CRF Variable": source_var,
                        "SDTM IG Target Raw": raw_target,
                        "Unparsed Token": token
                    })

    mapping_df = (
        pd.DataFrame(mapping_records)
        .drop_duplicates()
        .sort_values(by=["SDTM Domain", "SDTM Variable"])
        .reset_index(drop=True)
    ) if mapping_records else pd.DataFrame(columns=["SDTM Domain", "SDTM Variable"])

    detail_df = (
        pd.DataFrame(detail_records)
        .drop_duplicates()
        .sort_values(["SDTM Domain", "SDTM Variable", "CRF Dataset"])
        .reset_index(drop=True)
    ) if detail_records else pd.DataFrame()

    return mapping_df, detail_df, sheet_errors, unparsed_records
    # End=========================================================




# 整批CRF處理
def process_uploaded_excel(file_bytes, all_sheets):

    # Step1 context
    ctx = build_step1_context(file_bytes, all_sheets)

    soa_df = ctx["soa_df"]
    folder_df = ctx["folder_df"]
    domain_df_map = ctx["domain_df_map"]


    # SoA list
    soa_list_df = build_soa_visit_list(soa_df, folder_df)

    # SDTM mapping
    mapping_df, detail_df, mapping_errors, unparsed_records = build_sdtm_mapping(
        domain_df_map
    )

    # CT mapping
    ct_mapping_df, ct_mapping_sheet_errors = build_ct_mapping_seed(
        domain_df_map
    )


    return {
        "soa_list_df": soa_list_df,
        "mapping_df": mapping_df,
        "detail_df": detail_df,
        "unparsed_records": unparsed_records,
        "mapping_sheet_errors": mapping_errors,
        "ct_mapping_df": ct_mapping_df,
        "ct_mapping_sheet_errors": ct_mapping_sheet_errors,
        "available_sheets": ctx["available_sheets"],
        "missing_sheets": ctx["missing_sheets"],
        "sheet_errors": ctx["sheet_errors"]
    }
    # End=========================================================





# =================================================================================================================
# CT系列
# =================================================================================================================

# 抓CRF Option Displayed Value欄位
def find_option_displayed_value_column(columns):
    """
    優先抓 Option Displayed Value
    """
    priority_exact = [
        "OPTION DISPLAYED VALUE",
        "OPTION_DISPLAYED_VALUE",
        "OPTION DISPLAY VALUE",
        "DISPLAYED VALUE",
        "OPTION VALUE",
        "OPTIONS"
    ]

    normalized_map = {col: normalize_text(col) for col in columns}

    # 1. 先找完全一致
    for target in priority_exact:
        for col, norm_col in normalized_map.items():
            if norm_col == target:
                return col

    # 2. 再找包含 OPTION + DISPLAY + VALUE
    for col, norm_col in normalized_map.items():
        if "OPTION" in norm_col and "DISPLAY" in norm_col and "VALUE" in norm_col:
            return col

    # 3. 再退一步找 OPTION + VALUE
    for col, norm_col in normalized_map.items():
        if "OPTION" in norm_col and "VALUE" in norm_col:
            return col

    return None
    # End=========================================================


# 處理CRF Option Displayed Value
def split_option_displayed_values(value):
    """
    將 Option Displayed Value 拆成多個 option
    規則：
      - 只用分號 ; 和換行切
      - 不用逗號和斜線切
    """
    if pd.isna(value):
        return []

    text = str(value).strip()
    if not text:
        return []

    tokens = re.split(r"[;\n]+", text)

    out = []
    for token in tokens:
        token = str(token).strip()
        if token:
            out.append(token)

    return out
    # End=========================================================



# 處理CRF -> SDTM CT Mapping
def build_ct_mapping_seed(domain_df_map):
    """
    從已讀入的 CRF Domain DataFrames 建立 CT Mapping Seed

    輸出：
      - ct_mapping_df
      - ct_mapping_sheet_errors
    """

    seed_records = []
    ct_mapping_sheet_errors = []

    for sheet, df in domain_df_map.items():

        try:
            target_col = find_column(df.columns, ["SDTM", "TARGET"])
            source_var_col = find_source_variable_column(df.columns)
            option_col = find_option_displayed_value_column(df.columns)

            # 至少要有 SDTM Target 與 Source Variable
            if target_col is None or source_var_col is None:
                ct_mapping_sheet_errors.append(sheet)
                continue

            # 沒有 option 欄位，不算錯，但這張無法產 seed
            if option_col is None:
                continue

            for _, row in df.iterrows():

                raw_target = row.get(target_col, "")
                source_var = row.get(source_var_col, "")
                raw_option = row.get(option_col, "")

                source_var = "" if pd.isna(source_var) else str(source_var).strip()
                raw_target = "" if pd.isna(raw_target) else str(raw_target).strip()
                raw_option = "" if pd.isna(raw_option) else str(raw_option).strip()

                # 沒有 source var 就跳過
                if not source_var:
                    continue

                # 先 parse SDTM target
                parsed_records, _ = parse_sdtm_targets(raw_target)

                # 沒 parse 到 SDTM target，就不進 seed
                if not parsed_records:
                    continue

                # 拆 options
                option_tokens = split_option_displayed_values(raw_option)

                # 沒 option 的變數，不進 CT seed
                if not option_tokens:
                    continue

                for rec in parsed_records:
                    for opt in option_tokens:
                        seed_records.append({
                            "CRF Dataset": sheet,
                            "CRF Variable": source_var,
                            "SDTM Domain": rec["SDTM Domain"],
                            "SDTM Variable": rec["SDTM Variable"],
                            "Assign Value": rec["Assign Value"],
                            "SDTM IG Target Raw": raw_target,
                            "Option Displayed Value Raw": raw_option,
                            "Option Displayed Value": opt,
                            "Option Normalized": normalize_text(opt)
                        })

        except Exception:
            ct_mapping_sheet_errors.append(sheet)
            continue

    if seed_records:
        ct_mapping_df = (
            pd.DataFrame(seed_records)
            .drop_duplicates()
            .sort_values(
                by=[
                    "SDTM Domain",
                    "SDTM Variable",
                    "CRF Dataset",
                    "CRF Variable",
                    "Option Displayed Value"
                ]
            )
            .reset_index(drop=True)
        )
    else:
        ct_mapping_df = pd.DataFrame(columns=[
            "CRF Dataset",
            "CRF Variable",
            "SDTM Domain",
            "SDTM Variable",
            "Assign Value",
            "SDTM IG Target Raw",
            "Option Displayed Value Raw",
            "Option Displayed Value",
            "Option Normalized"
        ])

    return ct_mapping_df, sorted(list(set(ct_mapping_sheet_errors)))
    # End=========================================================



# =================================================================================================================
# 系統流程設定
# =================================================================================================================
# 判斷 Step 1 結果能不能重用（避免每次重跑）
def make_step1_cache_key(file_bytes):
    md5 = hashlib.md5(file_bytes).hexdigest()
    return md5
    # End=========================================================












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
       
        result = process_uploaded_excel(
            file_bytes=file_bytes,
            all_sheets=all_sheets
        )
        # 呼叫SoA
        soa_df = result["soa_list_df"]
        st.write(result)

        # Visit去重複 (供後續TV使用)
        unique_visit_df = (           
            soa_df
            .loc[
                (soa_df["Visit"].notna()) &
                (soa_df["Visit"].str.strip() != "") &
                (soa_df["CRF Dataset"] != soa_df["Abbreviation"])
            ]
            [["Abbreviation", "Visit", "Visit_order"]]
            .drop_duplicates()
            .sort_values("Visit_order")
            .reset_index(drop=True)
        )
        
        st.session_state["unique_visit_df"] = unique_visit_df
        st.write(unique_visit_df)
        
    
        
    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
