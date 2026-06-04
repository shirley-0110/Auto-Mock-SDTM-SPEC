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
from st_aggrid import AgGrid, GridOptionsBuilder

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



# 抓檔名 (Sponsor, Protocol)
def extract_protocol_no_from_filename(file_name):
    if not file_name:
        return ""

    name = os.path.splitext(file_name)[0].strip()

    parts = re.split(r"ecrf\s*schema", name, flags=re.IGNORECASE)
    prefix = parts[0].strip().strip("_") if parts else name

    tokens = [t.strip() for t in prefix.split("_") if t.strip()]

    sponsor = ""
    protocol = ""

    if len(tokens) >= 2:
        sponsor = tokens[0]
        protocol = tokens[1]
    elif len(tokens) == 1:
        protocol = tokens[0]

    return sponsor, protocol
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
# 匯入/整理Config
# =================================================================================================================
def load_domains_config(version):
    if version == "Version 3.3":
        path = "config/v33/domains.sas7bdat"
    else:
        path = "config/v34/domains.sas7bdat"

    if not HAS_PYREADSTAT:
        raise ImportError("目前環境尚未安裝 pyreadstat，請先在 requirements.txt 加入 pyreadstat")

    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到 config 檔：{path}")

    cfg_df, _ = pyreadstat.read_sas7bdat(path)
    cfg_df = normalize_columns(cfg_df)

    return cfg_df, path
    # End=========================================================


def standardize_domains_config(cfg_df):
    df = cfg_df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    rename_map = {
        "domain": "Dataset",
        "dlabel": "Dataset Label",
        "repeat": "Repeat",
        "refdata": "RefData",
        "structure": "Structure",
        "keyvars": "Key Variables",
        "keyseq": "KeySeq",
        "varnum": "VarNum",
        "name": "Variable",
        "label": "Variable Label",
        "type": "Data Type",
        "mandatory": "Mandatory",
        "role": "Role",
        "core": "Core",
        "ctcode": "CTcode",
        "class": "Class"
    }

    df = df.rename(columns=rename_map)

    if "Dataset" in df.columns:
        df["Dataset"] = df["Dataset"].astype(str).str.upper().str.strip()

    if "Variable" in df.columns:
        df["Variable"] = df["Variable"].astype(str).str.upper().str.strip()

    return df
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
        domain_df_map,
        st.session_state["var_to_ctcode"]
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
def build_ct_mapping_seed(domain_df_map, var_to_ctcode):
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

            for _, row in df.iterrows():
                try:
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

                    for rec in parsed_records:

                        sdtm_var = str(rec["SDTM Variable"]).strip().upper()
                        ctcode = var_to_ctcode.get(sdtm_var, "")
                        assign_val = rec.get("Assign Value", "")
                        assign_val = "" if pd.isna(assign_val) else str(assign_val).strip()


                        if not ctcode:
                            continue

                        # Assign Value 優先；否則用 option_tokens
                        if assign_val:
                            orival_candidates = [(assign_val, "")]
                        else:
                            if not option_tokens:
                                continue
                            orival_candidates = [(opt, opt) for opt in option_tokens]

                        for orival, option_displayed_value in orival_candidates:
                            orival = "" if pd.isna(orival) else str(orival).strip()

                            if not orival:
                                continue

                            seed_records.append({
                                "SDTM Domain": str(rec["SDTM Domain"]).strip(),
                                "SDTM Variable": sdtm_var,
                                "CTcode": ctcode,
                                "ORIVAL Normalized": normalize_text(orival)
                            })

                except Exception:
                    # 單列失敗不影響整張 sheet
                    continue

        except Exception:
            ct_mapping_sheet_errors.append(sheet)
            continue

    if seed_records:
        ct_mapping_df = (
            pd.DataFrame(seed_records)
            .drop_duplicates(
                subset=[
                    "SDTM Domain",
                    "SDTM Variable",
                    "CTcode",
                    "ORIVAL Normalized"
                ]
            )
            .sort_values(
                by=[
                    "SDTM Domain",
                    "SDTM Variable",
                    "CTcode",
                    "ORIVAL Normalized",
                ]
            )
            .reset_index(drop=True)
        )
    else:
        ct_mapping_df = pd.DataFrame(columns=[
            "SDTM Domain",
            "SDTM Variable",
            "CTcode",
            "ORIVAL Normalized"
        ])

    return ct_mapping_df, sorted(list(set(ct_mapping_sheet_errors)))
    # End=========================================================





def standardize_ct_mapping_dict(df):
    df = df.copy()
    df = normalize_columns(df)

    rename_map = {
        "CTCODE": "CTcode",
        "ORIVAL NORMALIZED": "ORIVAL Normalized",
        "CTVAL": "CTVAL",
        "ACTIVE": "Active",
        "NOTES": "Notes"
    }

    # 先把欄名 normalize 到大寫，再對映
    normalized_col_map = {c: normalize_text(c) for c in df.columns}
    new_cols = {}
    for original_col, norm_col in normalized_col_map.items():
        if norm_col in rename_map:
            new_cols[original_col] = rename_map[norm_col]

    df = df.rename(columns=new_cols)

    required_cols = ["CTcode", "ORIVAL Normalized", "CTVAL"]
    for c in required_cols:
        if c not in df.columns:
            raise ValueError(f"CT Mapping Dictionary 缺少必要欄位: {c}")

    df["CTcode"] = df["CTcode"].astype(str).str.strip().str.upper()
    df["ORIVAL Normalized"] = df["ORIVAL Normalized"].astype(str).str.strip().str.upper()
    df["CTVAL"] = df["CTVAL"].astype(str).str.strip()

    if "Active" in df.columns:
        df["Active"] = df["Active"].astype(str).str.strip().str.upper()
        df = df[df["Active"] != "N"]

    df = df[
        (df["CTcode"] != "") &
        (df["ORIVAL Normalized"] != "") &
        (df["CTVAL"] != "")
    ].drop_duplicates(subset=["CTcode", "ORIVAL Normalized"], keep="first")

    return df
    # End=========================================================




def build_ct_mapping(ct_seed_df, mapping_dict_df, ct_alias_df=None):
    """
    依 CTcode + ORIVAL Normalized 做 CT mapping

    Parameters
    ----------
    ct_seed_df : DataFrame
        來自 build_ct_mapping_seed() 的輸出
        必要欄位:
        - SDTM Domain
        - SDTM Variable
        - CTcode
        - ORIVAL Normalized

    mapping_dict_df : DataFrame
        開發者維護的 CT Mapping Dictionary
        必要欄位:
        - CTcode
        - ORIVAL Normalized
        - CTVAL

    ct_alias_df : DataFrame, optional
        CT 主檔整理出的 alias 表，用來 fallback
        建議欄位:
        - CTcode
        - ORIVAL Normalized
        - CTVAL
        - Alias Source   (optional)

    Returns
    -------
    matched_df : DataFrame
    unmatched_df : DataFrame
    """

    if ct_seed_df is None or ct_seed_df.empty:
        empty_cols = [
            "SDTM Domain", "SDTM Variable", "CTcode", "ORIVAL Normalized",
            "CTVAL", "Match Method"
        ]
        return pd.DataFrame(columns=empty_cols), pd.DataFrame(columns=empty_cols)

    seed = ct_seed_df.copy()
    seed = normalize_columns(seed)

    # 統一 seed 欄名（如果 normalize_columns 後有大小寫差異）
    seed_col_map = {c: normalize_text(c) for c in seed.columns}
    rename_seed = {}
    for original, norm in seed_col_map.items():
        if norm == "SDTM DOMAIN":
            rename_seed[original] = "SDTM Domain"
        elif norm == "SDTM VARIABLE":
            rename_seed[original] = "SDTM Variable"
        elif norm == "CTCODE":
            rename_seed[original] = "CTcode"
        elif norm == "ORIVAL NORMALIZED":
            rename_seed[original] = "ORIVAL Normalized"

    seed = seed.rename(columns=rename_seed)

    required_seed_cols = [
        "SDTM Domain", "SDTM Variable", "CTcode", "ORIVAL Normalized"
    ]
    for c in required_seed_cols:
        if c not in seed.columns:
            raise ValueError(f"CT Seed 缺少必要欄位: {c}")

    seed["CTcode"] = seed["CTcode"].astype(str).str.strip().str.upper()
    seed["ORIVAL Normalized"] = seed["ORIVAL Normalized"].astype(str).str.strip().str.upper()

    # 1) 主 mapping dictionary
    mapping_dict = standardize_ct_mapping_dict(mapping_dict_df)

    mapping_dict["CTcode"] = mapping_dict["CTcode"].str.upper().str.strip()
    mapping_dict["ORIVAL Normalized"] = mapping_dict["ORIVAL Normalized"].apply(normalize_text)

    matched = seed.merge(
        mapping_dict[["CTcode", "ORIVAL Normalized", "CTVAL"]],
        how="left",
        on=["CTcode", "ORIVAL Normalized"]
    )

    matched["Match Method"] = matched["CTVAL"].apply(
        lambda x: "DICT" if pd.notna(x) and str(x).strip() != "" else ""
    )

    # 2) fallback：CT alias exact match（optional）
    if ct_alias_df is not None and not ct_alias_df.empty:
        alias_df = ct_alias_df.copy()
        alias_df = normalize_columns(alias_df)

        alias_col_map = {c: normalize_text(c) for c in alias_df.columns}
        rename_alias = {}
        for original, norm in alias_col_map.items():
            if norm == "CTCODE":
                rename_alias[original] = "CTcode"
            elif norm == "ORIVAL NORMALIZED":
                rename_alias[original] = "ORIVAL Normalized"
            elif norm == "CTVAL":
                rename_alias[original] = "CTVAL"
            elif norm == "ALIAS SOURCE":
                rename_alias[original] = "Alias Source"

        alias_df = alias_df.rename(columns=rename_alias)

        required_alias_cols = ["CTcode", "ORIVAL Normalized", "CTVAL"]
        has_alias = all(c in alias_df.columns for c in required_alias_cols)

        if has_alias:
            alias_df["CTcode"] = alias_df["CTcode"].astype(str).str.strip().str.upper()
            alias_df["ORIVAL Normalized"] = alias_df["ORIVAL Normalized"].astype(str).str.strip().str.upper()
            alias_df["CTVAL"] = alias_df["CTVAL"].astype(str).str.strip()

            alias_df = alias_df.drop_duplicates(subset=["CTcode", "ORIVAL Normalized"], keep="first")

            need_fallback = matched["CTVAL"].isna() | (matched["CTVAL"].astype(str).str.strip() == "")
            fallback_seed = matched.loc[need_fallback, required_seed_cols].copy()

            if not fallback_seed.empty:
                fallback_hit = fallback_seed.merge(
                    alias_df[["CTcode", "ORIVAL Normalized", "CTVAL"]],
                    how="left",
                    on=["CTcode", "ORIVAL Normalized"]
                )

                for idx, row in fallback_hit.iterrows():
                    ctval = row["CTVAL"]
                    if pd.notna(ctval) and str(ctval).strip() != "":
                        original_idx = fallback_seed.index[idx]
                        matched.loc[original_idx, "CTVAL"] = ctval
                        matched.loc[original_idx, "Match Method"] = "ALIAS_EXACT"

    # 3) split matched / unmatched
    matched_mask = matched["CTVAL"].notna() & (matched["CTVAL"].astype(str).str.strip() != "")
    matched_df = matched.loc[matched_mask].copy()
    unmatched_df = matched.loc[~matched_mask].copy()

    # 4) 排序
    sort_cols = [c for c in [
        "SDTM Domain", "SDTM Variable", "CTcode", "ORIVAL Normalized"
    ] if c in matched_df.columns]

    if not matched_df.empty:
        matched_df = matched_df.sort_values(sort_cols).reset_index(drop=True)

    if not unmatched_df.empty:
        unmatched_df = unmatched_df.sort_values(sort_cols).reset_index(drop=True)

    return matched_df, unmatched_df
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


uploaded_file = st.file_uploader("請上傳 eCRF Schema Excel", type=["xlsx", "xls"])

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
        # 前置作業
        # -------------------------------------------------
        # Sponsor_Protocol
        st.markdown("#### Basic Information")

        sponsor, default_protocol_no = extract_protocol_no_from_filename(uploaded_file.name)

        col1, col2 = st.columns(2)

        with col1:
            protocol_no = st.text_input(
                "Protocol No",
                value=default_protocol_no,
                key="protocol_no"
            )

        with col2:
            protocol_title = st.text_input(
                "Protocol Title",
                value="",
                key="protocol_title"
            )

        # Version Control
        st.markdown("#### Version Control")

        r1_c1, r1_c2 = st.columns(2)

        with r1_c1:
            version = st.selectbox(
                "SDTM IG",
                ["Version 3.4", "Version 3.3"],
                key="sdtm_version_selector"
            )

        with r1_c2:
            sdtm_ct = st.text_input("SDTM CT", value="", key="sdtm_ct")
        
        meddra_version = st.text_input("MedDRA", value="", key="meddra_version")

        r3_c1, r3_c2 = st.columns(2)

        with r3_c1:
            cm_dictionary = st.selectbox(
                "CM 字典",
                ["WHODrug Global B3", "WHO ATC/DDD"],
                key="cm_dictionary"
            )

        with r3_c2:
            cm_version = st.text_input("CM 版本", value="", key="cm_version")

        
        r4_c1, r4_c2, r4_c3 = st.columns(3)

        with r4_c1:
            snomed_version = st.text_input("SNOMED", value="", key="snomed_version")

        with r4_c2:
            unii_version = st.text_input("UNII", value="", key="unii_version")

        with r4_c3:
            medrt_version = st.text_input("MED-RT", value="", key="medrt_version")


        # Load Config
        if (
            not st.session_state.get("config_loaded")
            or st.session_state.get("config_version") != version
        ):

            raw_cfg_df, cfg_path = load_domains_config(version)

            cfg_df = standardize_domains_config(raw_cfg_df)
    
            # 存 config
            st.session_state["config_df"] = cfg_df

            # 建 mapping（CT mapping會用）
            st.session_state["var_to_ctcode"] = dict(
                zip(cfg_df["Variable"], cfg_df["CTcode"])
            )

            # 記錄版本
            st.session_state["config_version"] = version
            st.session_state["config_loaded"] = True
        

        # Load CT Term Mapping Dictionary
        ct_mapping_dict_df = None

        dict_path = "config/CT Term Mapping Dictionary.xlsx"

        if os.path.exists(dict_path):
            try:
                ct_mapping_dict_df = pd.read_excel(dict_path)
                ct_mapping_dict_df = normalize_columns(ct_mapping_dict_df)

                st.session_state["ct_mapping_dict_df"] = ct_mapping_dict_df

                # st.success("CT Mapping Dictionary 載入成功")
            except Exception as e:
                st.warning(f"CT Mapping Dictionary 載入失敗：{e}")
        else:
            st.warning("找不到 CT Mapping Dictionary（目前只使用 seed）")


        
        # -------------------------------------------------
        # Step 1：CRF → SDTM Mapping
        # -------------------------------------------------
        st.markdown("## Step 1｜CRF → SDTM Mapping")

        step1_cache_key = make_step1_cache_key(file_bytes)

        if (
            st.session_state.get("step1_cache_key") == step1_cache_key
            and st.session_state.get("step1_result") is not None
        ):
            result = st.session_state["step1_result"]
        else:
            result = process_uploaded_excel(
                file_bytes=file_bytes,
                all_sheets=all_sheets
            )
            st.session_state["step1_cache_key"] = step1_cache_key
            st.session_state["step1_result"] = result


        # 呼叫SoA
        soa_df = result["soa_list_df"]

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
        # st.write(unique_visit_df)

        
        missing_sheets = result["missing_sheets"]
        mapping_df = result["mapping_df"]
        detail_df = result["detail_df"]
        sheet_errors = result["sheet_errors"]
        unparsed_records = result["unparsed_records"]
        ct_mapping_df = result.get("ct_mapping_df", pd.DataFrame())
        ct_mapping_sheet_errors = result.get("ct_mapping_sheet_errors", [])




        # SDTM Varialbe Mapping (Summary by Domain）
        st.markdown("### 📊 SDTM Variable Mapping")
        st.markdown("#### - Summary by Domain")
        
        if mapping_df.empty:
            st.warning("目前沒有從 CRF Sheet 抓到可解析的 SDTM Domain / Variable")
        else:
            summary_df = (
                mapping_df
                .groupby("SDTM Domain")["SDTM Variable"]
                .apply(lambda x: sorted(set(x)))
                .reset_index()
            )

            summary_df["Variable Count"] = summary_df["SDTM Variable"].apply(len)
            summary_df["Variables"] = summary_df["SDTM Variable"].apply(lambda x: "; ".join(x))

            st.dataframe(summary_df[["SDTM Domain", "Variable Count", "Variables"]], use_container_width=True)


        # Detail（CRF → SDTM）
        st.markdown("#### - Detail")
                
        if detail_df.empty:
            st.info("目前沒有可顯示的明細")
        else:          
            sorted_detail_df = detail_df.sort_values(
                by=["SDTM Domain", "SDTM Variable", "CRF Dataset", "CRF Variable"],
                ascending=[True, True, True, True]
            ).reset_index(drop=True)

            st.dataframe(sorted_detail_df, use_container_width=True)



        # CT Seed
        st.markdown("### 🧩 CT Mapping Seed (Option level)")
        if ct_mapping_df.empty:
            st.info("目前沒有 CT Mapping Seed")
        else:
            st.dataframe(ct_mapping_df, use_container_width=True)




        # CT Mapping Result
        st.markdown("### 🧩 CT Mapping Result")

        # 確保 dictionary 已載入
        if "ct_mapping_dict_df" in st.session_state:

            mapping_dict_df = st.session_state["ct_mapping_dict_df"]

            matched_ct_df, unmatched_ct_df = build_ct_mapping(
                ct_mapping_df,
                mapping_dict_df
            )

            tab1, tab2 = st.tabs(["✅ Matched", "❌ Unmatched"])

            # -------------------------------------------------
            # ✅ Matched
            # -------------------------------------------------
            with tab1:

                if matched_ct_df.empty:
                    st.info("目前沒有 matched CT mapping")
                else:
                    display_cols = [
                        "SDTM Domain",
                        "SDTM Variable",
                        "CTcode",
                        "ORIVAL",
                        "ORIVAL Normalized",
                        "CTVAL",
                        "Match Method"
                    ]

                    display_cols = [c for c in display_cols if c in matched_ct_df.columns]

                    st.dataframe(
                        matched_ct_df[display_cols],
                        use_container_width=True
                    )


            # -------------------------------------------------
            # ❌ Unmatched（超重要）
            # -------------------------------------------------
            with tab2:
    
                if unmatched_ct_df.empty:
                    st.success("🎉 所有 ORIVAL 都已成功 mapping")
                else:

                    st.warning("以下 ORIVAL 尚未對應 CTVAL，建議加入 CT Mapping Dictionary")

                    display_cols = [
                        "SDTM Domain",
                        "SDTM Variable",
                        "CTcode",
                        "ORIVAL",
                        "ORIVAL Normalized"
                    ]

                    display_cols = [c for c in display_cols if c in unmatched_ct_df.columns]

                    st.dataframe(
                        unmatched_ct_df[display_cols],
                        use_container_width=True
                    )

                    # ✅ ✅ ✅ Bonus：開發者直接匯出用
                    st.download_button(
                        label="⬇️ 下載 Unmatched CT（用來補 mapping）",
                        data=unmatched_ct_df.to_csv(index=False),
                        file_name="ct_mapping_unmatched.csv",
                        mime="text/csv"
                    )

        else:
            st.info("尚未載入 CT Mapping Dictionary")



        
        # 錯誤 / Debug
        st.markdown("### ⚠️ Debug / Error 檢查")

        if missing_sheets:
            st.warning(f"SoA 有出現的 Form OID，但 Excel 沒有對應的 Sheets: {missing_sheets}")
        
        if sheet_errors:
            st.warning(f"無法處理的 Sheets (Header偵測失敗): {sorted(set(sheet_errors))}")
        
        if unparsed_records:
            st.markdown("#### 無法解析的 SDTM IG Target")
            st.dataframe(pd.DataFrame(unparsed_records), use_container_width=True)
        
        if ct_mapping_sheet_errors:
            st.warning(f"CT Mapping 無法處理的 Sheets：{sorted(set(ct_mapping_sheet_errors))}")

    
        
    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
