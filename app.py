import streamlit as st
import pandas as pd
import re
import os
import hashlib

from io import BytesIO

# Step 2 用到 sas7bdat
try:
    import pyreadstat
    HAS_PYREADSTAT = True
except Exception:
    HAS_PYREADSTAT = False

st.set_page_config(page_title="Auto SDTM SPEC", layout="wide")
st.title("Auto SDTM SPEC")


# =========================================================
# 基本工具函式
# =========================================================
def normalize_text(x):
    if pd.isna(x):
        return ""
    x = str(x)
    x = x.replace("\n", " ")
    x = x.replace("\r", " ")
    x = x.replace("\xa0", " ")
    x = re.sub(r"\s+", " ", x)
    return x.strip().upper()


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


def row_contains_keywords(row_values, keyword_groups):
    cells = [normalize_text(v) for v in row_values]

    for cell in cells:
        for group in keyword_groups:
            if all(k.upper() in cell for k in group):
                return True
    return False


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


def find_column(columns, required_keywords):
    for col in columns:
        upper_col = normalize_text(col)
        if all(k.upper() in upper_col for k in required_keywords):
            return col
    return None


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



def find_option_displayed_value_column(columns):
    """
    抓 CRF schema 裡的 option 顯示值欄位
    常見名稱例如：
      - Option Displayed Value
      - Displayed Value
      - Option Label
      - Decode
    """
    priority_exact = [
        "OPTION DISPLAYED VALUE",
        "OPTION DISPLAY VALUE",
        "DISPLAYED VALUE",
        "OPTION LABEL",
        "OPTION TEXT",
        "DECODE",
        "CODELIST DISPLAYED VALUE"
    ]

    normalized_map = {col: normalize_text(col) for col in columns}

    for target in priority_exact:
        for col, norm_col in normalized_map.items():
            if norm_col == target:
                return col

    for col, norm_col in normalized_map.items():
        if "OPTION" in norm_col and "DISPLAY" in norm_col and "VALUE" in norm_col:
            return col

    for col, norm_col in normalized_map.items():
        if "DISPLAYED" in norm_col and "VALUE" in norm_col:
            return col

    for col, norm_col in normalized_map.items():
        if "OPTION" in norm_col and "LABEL" in norm_col:
            return col

    return None


# =========================================================
# SoA：抓 CRF domain / sheet
# =========================================================
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


# =========================================================
# SDTM IG Target parsing
# =========================================================
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


def build_sdtm_mapping(file_bytes, selected_crf_sheets, common_domain_header=None):
    """
    common_domain_header:
      所有 Domain Sheet 共用 header row（Excel 1-based）
    """
    mapping_records = []
    detail_records = []
    sheet_errors = []
    unparsed_records = []

    for sheet in selected_crf_sheets:
        try:
            df, _ = read_sheet_with_detected_header(
                file_bytes=file_bytes,
                sheet_name=sheet,
                keyword_groups=[["SDTM", "TARGET"]],
                manual_header_row_excel=common_domain_header
            )
        except Exception:
            sheet_errors.append(sheet)
            continue

        target_col = find_column(df.columns, ["SDTM", "TARGET"])
        if target_col is None:
            sheet_errors.append(sheet)
            continue

        source_var_col = find_source_variable_column(df.columns)

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
                    "Source CRF Sheet": sheet,
                    "Source CRF Variable": source_var,
                    "SDTM Domain": rec["SDTM Domain"],
                    "SDTM Variable": rec["SDTM Variable"],
                    "Assign Value": rec["Assign Value"],
                    "SDTM IG Target Raw": raw_target
                })

            for token in unparsed_tokens:
                if str(token).strip():
                    unparsed_records.append({
                        "Source CRF Sheet": sheet,
                        "Source CRF Variable": source_var,
                        "SDTM IG Target Raw": raw_target,
                        "Unparsed Token": token
                    })

    if mapping_records:
        mapping_df = (
            pd.DataFrame(mapping_records)
            .drop_duplicates()
            .sort_values(by=["SDTM Domain", "SDTM Variable"])
            .reset_index(drop=True)
        )
    else:
        mapping_df = pd.DataFrame(columns=["SDTM Domain", "SDTM Variable"])

    if detail_records:
        detail_df = (
            pd.DataFrame(detail_records)
            .drop_duplicates()
            .sort_values(
                by=[
                    "SDTM Domain",
                    "SDTM Variable",
                    "Source CRF Sheet",
                    "Source CRF Variable",
                    "Assign Value"
                ]
            )
            .reset_index(drop=True)
        )
    else:
        detail_df = pd.DataFrame(columns=[
            "Source CRF Sheet",
            "Source CRF Variable",
            "SDTM Domain",
            "SDTM Variable",
            "Assign Value",
            "SDTM IG Target Raw"
        ])

    return mapping_df, detail_df, sheet_errors, unparsed_records


def summarize_sdtm_mapping(mapping_df):
    if mapping_df.empty:
        return pd.DataFrame(columns=["SDTM Domain", "Variable Count", "Variables"])

    summary_df = (
        mapping_df.groupby("SDTM Domain")["SDTM Variable"]
        .apply(lambda x: sorted(set(x)))
        .reset_index()
    )

    summary_df["Variable Count"] = summary_df["SDTM Variable"].apply(len)
    summary_df["Variables"] = summary_df["SDTM Variable"].apply(lambda x: "; ".join(x))

    return summary_df[["SDTM Domain", "Variable Count", "Variables"]]





def build_ct_mapping_seed(file_bytes, selected_crf_sheets, common_domain_header=None):
    """
    從各 domain sheet 抽取 CT mapping seed
    輸出欄位：
      - Dataset
      - Variable
      - Source CRF Sheet
      - Source CRF Variable
      - Assign Value
      - Option Displayed Value
      - Normalized Option
      - Suggested CT Term
      - Match Status

    說明：
      - 只要該列有 Option Displayed Value，且可解析出 SDTM target，就納入
      - 先不做 CT merge，只做 seed list
    """
    records = []
    sheet_errors = []

    for sheet in selected_crf_sheets:
        try:
            df, _ = read_sheet_with_detected_header(
                file_bytes=file_bytes,
                sheet_name=sheet,
                keyword_groups=[["SDTM", "TARGET"]],
                manual_header_row_excel=common_domain_header
            )
        except Exception:
            sheet_errors.append(sheet)
            continue

        target_col = find_column(df.columns, ["SDTM", "TARGET"])
        if target_col is None:
            continue

        source_var_col = find_source_variable_column(df.columns)
        option_display_col = find_option_displayed_value_column(df.columns)

        # 沒有 option 顯示值欄位，這個 sheet 就先略過
        if option_display_col is None:
            continue

        for _, row in df.iterrows():
            raw_target = row.get(target_col, "")
            source_var = row.get(source_var_col, "") if source_var_col is not None else ""
            option_display = row.get(option_display_col, "")

            if pd.isna(option_display) or str(option_display).strip() == "":
                continue

            parsed_records, _ = parse_sdtm_targets(raw_target)

            if not parsed_records:
                continue

            for rec in parsed_records:
                option_text = str(option_display).strip()
                norm_option = re.sub(r"\s+", " ", option_text.upper())

                records.append({
                    "Dataset": rec["SDTM Domain"],
                    "Variable": rec["SDTM Variable"],
                    "Source CRF Sheet": sheet,
                    "Source CRF Variable": source_var,
                    "Assign Value": rec["Assign Value"],
                    "Option Displayed Value": option_text,
                    "Normalized Option": norm_option,
                    "Suggested CT Term": "",
                    "Match Status": "NEW"
                })

    if records:
        out_df = pd.DataFrame(records).drop_duplicates().sort_values(
            by=["Dataset", "Variable", "Source CRF Sheet", "Source CRF Variable", "Option Displayed Value"]
        ).reset_index(drop=True)
    else:
        out_df = pd.DataFrame(columns=[
            "Dataset", "Variable", "Source CRF Sheet", "Source CRF Variable",
            "Assign Value", "Option Displayed Value", "Normalized Option",
            "Suggested CT Term", "Match Status"
        ])

    return out_df, sorted(set(sheet_errors))






# =========================================================
# Step 2：domains.sas7bdat (單一 config 檔)
# =========================================================
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
        "ctcode": "Codelist",
        "class": "Class"
    }

    df = df.rename(columns=rename_map)

    if "Dataset" in df.columns:
        df["Dataset"] = df["Dataset"].astype(str).str.upper().str.strip()

    if "Variable" in df.columns:
        df["Variable"] = df["Variable"].astype(str).str.upper().str.strip()

    return df


def expand_suppqual_to_supp_datasets(config_df, detected_datasets):
    if config_df.empty:
        return config_df.copy()

    cfg = config_df.copy()

    if "Dataset" not in cfg.columns:
        return cfg

    suppqual_rows = cfg[cfg["Dataset"] == "SUPPQUAL"].copy()
    if suppqual_rows.empty:
        return cfg

    detected_supp = [ds for ds in detected_datasets if str(ds).upper().startswith("SUPP")]

    if not detected_supp:
        return cfg

    expanded_rows = [cfg]

    for ds in detected_supp:
        if ds == "SUPPQUAL":
            continue

        dup = suppqual_rows.copy()
        dup["Dataset"] = ds

        base_domain = ds[4:]  # SUPPAE -> AE
        dup["Dataset Label"] = f"Supplemental Qualifiers for {base_domain}"

        expanded_rows.append(dup)

    expanded_cfg = pd.concat(expanded_rows, ignore_index=True)
    expanded_cfg = expanded_cfg.drop_duplicates()

    return expanded_cfg.reset_index(drop=True)


def is_send_only(core_value):
    text = normalize_text(core_value)
    return "SEND" in text and "ONLY" in text


def should_keep_non_crf(row):
    variable = normalize_text(row.get("Variable", ""))
    core = normalize_text(row.get("Core", ""))
    return (core in ["REQUIRED", "EXPECTED"]) or (variable == "EPOCH")


def get_non_crf_from_config(detail_df, config_df):
    """
    non-CRF = config 裡所有 variables - Step 1 mapping 已有的 variables
    保留條件：
      - Core = Required 或 Expected
      - 或 Variable = EPOCH
    並排除 SEND Only
    """
    if detail_df.empty or config_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment", "Core", "VarNum"
        ])

    crf_pairs = set(
        zip(
            detail_df["SDTM Domain"].astype(str).str.upper(),
            detail_df["SDTM Variable"].astype(str).str.upper()
        )
    )

    detected_datasets = set(detail_df["SDTM Domain"].astype(str).str.upper())

    cfg = config_df.copy()
    cfg = cfg[
        (cfg["Dataset"].astype(str).str.strip() != "") &
        (cfg["Variable"].astype(str).str.strip() != "")
    ].copy()

    cfg = cfg[cfg["Dataset"].isin(detected_datasets)]

    if "Core" in cfg.columns:
        cfg = cfg[~cfg["Core"].apply(is_send_only)].copy()

    cfg["pair"] = list(zip(cfg["Dataset"], cfg["Variable"]))
    non_crf = cfg[~cfg["pair"].isin(crf_pairs)].copy()

    if "pair" in non_crf.columns:
        non_crf = non_crf.drop(columns=["pair"])

    non_crf = non_crf[non_crf.apply(should_keep_non_crf, axis=1)].copy()

    non_crf["Label"] = non_crf.get("Variable Label", "")
    non_crf["Origin"] = non_crf.get("Origin", "")
    non_crf["Source"] = non_crf.get("Source", "")
    non_crf["Pages"] = ""
    non_crf["Method"] = non_crf.get("Method", "")
    non_crf["Comment"] = non_crf.get("Comment", "")
    non_crf["Core"] = non_crf.get("Core", "")
    non_crf["IsCRFVariable"] = False

    out = non_crf[[
        "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment", "Core", "VarNum",
        "IsCRFVariable"
    ]].drop_duplicates(subset=["Dataset", "Variable"])

    return out.reset_index(drop=True)





def enrich_crf_variables_with_config(detail_df, config_df):
    """
    CRF 收集到的都保留
    並排除 SEND Only
    同一個 Dataset+Variable 只保留一筆

    規則：
      - 有 Assign Value -> Origin=Assigned
      - 其他 CRF 收集來的 -> Origin=Collected, Source=Investigator
      - 來自 CRF 的變數一律標記 IsCRFVariable=True
    """
    if detail_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment", "Core", "VarNum",
            "IsCRFVariable"
        ])

    detected_datasets = sorted(detail_df["SDTM Domain"].astype(str).str.upper().unique())
    expanded_cfg = expand_suppqual_to_supp_datasets(config_df, detected_datasets)

    if "Core" in expanded_cfg.columns:
        expanded_cfg = expanded_cfg[~expanded_cfg["Core"].apply(is_send_only)].copy()

    crf_df = detail_df.copy()
    crf_df["Dataset"] = crf_df["SDTM Domain"].astype(str).str.upper()
    crf_df["Variable"] = crf_df["SDTM Variable"].astype(str).str.upper()

    def derive_origin(row):
        if str(row.get("Assign Value", "")).strip() != "":
            return "Assigned"
        return "Collected"

    def derive_source(row):
        if str(row.get("Assign Value", "")).strip() != "":
            return ""
        return "Investigator"

    crf_df["Origin"] = crf_df.apply(derive_origin, axis=1)
    crf_df["Source"] = crf_df.apply(derive_source, axis=1)
    crf_df["Pages"] = ""
    crf_df["Method"] = ""
    crf_df["Comment"] = ""
    crf_df["IsCRFVariable"] = True

    cfg_keep_cols = [
        c for c in [
            "Dataset", "Variable", "Variable Label", "Data Type", "Codelist",
            "Core", "VarNum"
        ] if c in expanded_cfg.columns
    ]

    cfg = expanded_cfg[cfg_keep_cols].drop_duplicates() if cfg_keep_cols else pd.DataFrame(columns=["Dataset", "Variable"])

    merged = crf_df.merge(
        cfg,
        on=["Dataset", "Variable"],
        how="left"
    )

    merged["Label"] = merged.get("Variable Label", "")
    if "Core" not in merged.columns:
        merged["Core"] = ""
    if "VarNum" not in merged.columns:
        merged["VarNum"] = ""

    def join_unique(series):
        vals = [str(x).strip() for x in series if str(x).strip() not in ["", "nan", "None"]]
        vals = list(dict.fromkeys(vals))
        return "; ".join(vals)

    grouped = merged.groupby(["Dataset", "Variable"], dropna=False).agg({
        "Label": "first",
        "Data Type": "first",
        "Codelist": "first",
        "Origin": lambda s: "Assigned" if "Assigned" in list(s) else "Collected",
        "Source": lambda s: "Investigator" if "Investigator" in list(s) else "",
        "Pages": "first",
        "Method": "first",
        "Comment": join_unique,
        "Core": "first",
        "VarNum": "first",
        "IsCRFVariable": "first"
    }).reset_index()

    return grouped[[
        "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment", "Core", "VarNum",
        "IsCRFVariable"
    ]]




def normalize_data_type_by_config(raw_type, variable_name=""):
    """
    規則：
      - config type=1 -> integer
      - config type=2 -> text
      - 若 variable 以 DTC / STDTC / ENDTC 結尾 -> datetime
    """
    var = normalize_text(variable_name)

    if var.endswith("STDTC") or var.endswith("ENDTC") or var.endswith("DTC"):
        return "datetime"

    text = normalize_text(raw_type)

    if text in ["1", "1.0", "INTEGER", "INT", "NUMERIC"]:
        return "integer"
    if text in ["2", "2.0", "TEXT", "CHAR", "STRING"]:
        return "text"

    raw = str(raw_type).strip()
    if raw.lower() in ["nan", "none"]:
        return ""
    return raw


def build_config_variable_lookup(config_df):
    """
    建立 (Dataset, Variable) -> row dict lookup
    """
    lookup = {}

    if config_df.empty:
        return lookup

    temp = config_df.copy()

    if "Dataset" not in temp.columns or "Variable" not in temp.columns:
        return lookup

    for _, row in temp.iterrows():
        ds = str(row.get("Dataset", "")).strip().upper()
        var = str(row.get("Variable", "")).strip().upper()
        if ds and var:
            lookup[(ds, var)] = row.to_dict()

    return lookup



def build_variable_row_from_config(dataset, variable, cfg_lookup):
    """
    只在 config 找得到時才建立 row；
    找不到就回傳 None
    Auto-added variable 不加 comment
    """
    key = (str(dataset).upper(), str(variable).upper())

    if key not in cfg_lookup:
        return None

    meta = cfg_lookup[key]

    row = {
        "Dataset": str(meta.get("Dataset", dataset)).upper(),
        "Variable": str(meta.get("Variable", variable)).upper(),
        "Label": meta.get("Variable Label", ""),
        "Data Type": normalize_data_type_by_config(meta.get("Data Type", ""), variable),
        "Codelist": meta.get("Codelist", ""),
        "Origin": "Derived",
        "Source": "",
        "Pages": "",
        "Method": "",
        "Comment": "",
        "Core": meta.get("Core", ""),
        "VarNum": meta.get("VarNum", ""),
        "IsCRFVariable": False
    }

    return row



def append_required_partner_variables(final_df, config_df):
    """
    規則（僅在 config 有對應 variable 時才補）：
      1. --DTC    -> --DY
      2. --STDTC  -> --STDY
      3. --ENDTC  -> --ENDY
      4. --STRTPT -> --STTPT
      5. --ENRTPT -> --ENTPT
      6. VISITNUM -> VISIT, VISITDY
      7. --TPT    -> --TPTNUM
      8. --ORRES  -> --STRESC, --STRESN, --STAT
      9. --ORRESU -> --STRESU

    注意：
      - 若 config 沒有該變數，則不加
      - Auto-added variable 不加 Comment
    """
    if final_df.empty:
        return final_df

    out = final_df.copy()
    cfg_lookup = build_config_variable_lookup(config_df)

    existing_pairs = set(
        zip(
            out["Dataset"].astype(str).str.upper(),
            out["Variable"].astype(str).str.upper()
        )
    )

    new_rows = []

    for dataset, grp in out.groupby("Dataset", dropna=False):
        ds = str(dataset).upper()
        vars_in_ds = set(grp["Variable"].astype(str).str.upper())

        # =================================================
        # 無條件保留變數（domain-specific required variables）
        # =================================================

        # ---------- CO ----------
        if ds == "CO":
            for tgt_var in ["RDOMAIN", "IDVAR", "IDVARVAL", "COREF", "COEVAL"]:
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))

        # ---------- DS ----------
        if ds == "DS":
            for tgt_var in ["DSDTC", "DSDY"]:
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))


        # Rule A: exact VISITNUM -> VISIT, VISITDY
        if "VISITNUM" in vars_in_ds:
            for tgt_var in ["VISIT", "VISITDY"]:
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))
                        vars_in_ds.add(tgt_var)

        # Rule B: suffix-based
        for src_var in list(vars_in_ds):
            # --STDTC -> --STDY
            if src_var.endswith("STDTC"):
                tgt_var = src_var[:-5] + "STDY"
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))

            # --ENDTC -> --ENDY
            if src_var.endswith("ENDTC"):
                tgt_var = src_var[:-5] + "ENDY"
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))

            # --DTC -> --DY
            if src_var.endswith("DTC") and not src_var.endswith("STDTC") and not src_var.endswith("ENDTC"):
                tgt_var = src_var[:-3] + "DY"
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))

            # --STRTPT -> --STTPT
            if src_var.endswith("STRTPT"):
                tgt_var = src_var[:-6] + "STTPT"
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))

            # --ENRTPT -> --ENTPT
            if src_var.endswith("ENRTPT"):
                tgt_var = src_var[:-6] + "ENTPT"
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))

            # --TPT -> --TPTNUM
            if (
                src_var.endswith("TPT")
                and not src_var.endswith("STTPT")
                and not src_var.endswith("ENTPT")
            ):
                tgt_var = src_var + "NUM"
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))

            # --ORRES -> --STRESC, --STRESN, --STAT
            if src_var.endswith("ORRES"):
                base = src_var[:-5]
                for tgt_var in [base + "STRESC", base + "STRESN", base + "STAT"]:
                    if (ds, tgt_var) not in existing_pairs:
                        row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                        if row is not None:
                            new_rows.append(row)
                            existing_pairs.add((ds, tgt_var))

            # --ORRESU -> --STRESU
            if src_var.endswith("ORRESU"):
                base = src_var[:-6]
                tgt_var = base + "STRESU"
                if (ds, tgt_var) not in existing_pairs:
                    row = build_variable_row_from_config(ds, tgt_var, cfg_lookup)
                    if row is not None:
                        new_rows.append(row)
                        existing_pairs.add((ds, tgt_var))

    if new_rows:
        out = pd.concat([out, pd.DataFrame(new_rows)], ignore_index=True)

    return out


def build_trial_design_variables_spec(config_df):
    """
    TA / TE / TI / TS / TV 變數跟 config 比對，
    帶入 Core / VarNum / Codelist / Label / Data Type
    Comment 一律空白
    """
    defs = get_trial_design_definitions()
    cfg_lookup = build_config_variable_lookup(config_df)

    rows = []

    for domain in ["TA", "TE", "TI", "TS", "TV"]:
        for var, fallback_label, fallback_dtype in defs[domain]["variables"]:
            key = (domain, var)

            if key in cfg_lookup:
                meta = cfg_lookup[key]
                row = {
                    "Dataset": domain,
                    "Variable": var,
                    "Label": meta.get("Variable Label", fallback_label),
                    "Data Type": normalize_data_type_by_config(meta.get("Data Type", fallback_dtype), var),
                    "Codelist": meta.get("Codelist", ""),
                    "Origin": "Protocol",
                    "Source": "Sponsor",
                    "Pages": "",
                    "Method": "",
                    "Comment": "",
                    "Core": meta.get("Core", ""),
                    "VarNum": meta.get("VarNum", ""),
                    "IsCRFVariable": False
                }
            else:
                row = {
                    "Dataset": domain,
                    "Variable": var,
                    "Label": fallback_label,
                    "Data Type": normalize_data_type_by_config(fallback_dtype, var),
                    "Codelist": "",
                    "Origin": "Protocol",
                    "Source": "Sponsor",
                    "Pages": "",
                    "Method": "",
                    "Comment": "",
                    "Core": "",
                    "VarNum": "",
                    "IsCRFVariable": False
                }

            rows.append(row)

    return pd.DataFrame(rows)





def apply_origin_source_method_overrides(df):
    """
    最終修正 Origin / Source / Method / 特定 Codelist
    注意：
      - 只有 IsCRFVariable=False 的列才允許被 override
      - CRF 來的變數一律保留原本狀態，不被 non-CRF 規則覆蓋
    """
    if df.empty:
        return df

    out = df.copy()

    for c in ["Dataset", "Variable", "Origin", "Source", "Method", "Codelist", "IsCRFVariable"]:
        if c not in out.columns:
            out[c] = ""

    ds = out["Dataset"].astype(str).str.upper()
    var = out["Variable"].astype(str).str.upper()

    # True 表示來自 CRF -> 不可 override
    is_crf = out["IsCRFVariable"].fillna(False).astype(bool)

    # helper: 只對 non-CRF 套規則
    def non_crf_mask(base_mask):
        return base_mask & (~is_crf)

    # -------------------------------------------------
    # Core identifiers
    # -------------------------------------------------
    mask = non_crf_mask(var == "STUDYID")
    out.loc[mask, "Origin"] = "Protocol"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = ""

    mask = non_crf_mask(var == "DOMAIN")
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = ""

    mask = non_crf_mask(var == "USUBJID")
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = "Concatenation of STUDYID-SITEID-SUBJID"

    # -------------------------------------------------
    # Sequence
    # -------------------------------------------------
    mask = non_crf_mask((var.str.endswith("SEQ")) & (var != "TSSEQ"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = (
        "Equal to sequential number identifying records within each USUBJID "
        "which sorted by key variables in the domain"
    )

    mask = non_crf_mask(var == "TSSEQ")
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = (
        "Equal to sequential number identifying records within each TSPARMCD in the domain"
    )

    # -------------------------------------------------
    # EPOCH
    # -------------------------------------------------
    mask = non_crf_mask(var == "EPOCH")
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = ""

    # TA.EPOCH 例外
    mask = non_crf_mask((ds == "TA") & (var == "EPOCH"))
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = ""

    # -------------------------------------------------
    # AE dictionary variables
    # -------------------------------------------------
    ae_dict_vars = {
        "AELLT", "AELLTCD", "AEDECOD", "AEPTCD",
        "AEHLT", "AEHLTCD", "AEHLGT", "AEHLGTCD",
        "AEBODSYS", "AEBDSYCD", "AESOC", "AESOCCD"
    }
    mask = non_crf_mask(var.isin(ae_dict_vars))
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Vendor"


    # -------------------------------------------------
    # VISIT / VISITNUM / VISITDY
    # -------------------------------------------------
    mask = non_crf_mask(var == "VISITNUM")
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = ""
    
    mask = non_crf_mask(var == "VISIT")
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = ""

    mask = non_crf_mask(var == "VISITDY")
    out.loc[mask, "Origin"] = "Protocol"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = ""

    # -------------------------------------------------
    # STTPT / ENTPT
    # -------------------------------------------------
    mask = non_crf_mask(var.str.endswith("STTPT") | var.str.endswith("ENTPT"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"

    # -------------------------------------------------
    # STRES / STAT
    # -------------------------------------------------
    mask = non_crf_mask(var.str.endswith("STRESC"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = var[mask].str.replace("STRESC", "ORRES", regex=False).apply(
        lambda x: f"Equal to {x}"
    )

    mask = non_crf_mask(var.str.endswith("STRESN"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = var[mask].str.replace("STRESN", "STRESC", regex=False).apply(
        lambda x: f"Equal to numeric value of {x} if {x} contains numeric data"
    )

    mask = non_crf_mask(var.str.endswith("STRESU"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = var[mask].str.replace("STRESU", "ORRESU", regex=False).apply(
        lambda x: f"Equal to {x}"
    )

    mask = non_crf_mask(var.str.endswith("STAT"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = var[mask].str.replace("STAT", "ORRES", regex=False).apply(
        lambda x: f'Equal to "NOT DONE" if {x} is null'
    )

    # -------------------------------------------------
    # DY / STDY / ENDY
    # -------------------------------------------------
    def build_dy_method(v):
        if v.endswith("STDY"):
            src = v[:-4] + "STDTC"
        elif v.endswith("ENDY"):
            src = v[:-4] + "ENDTC"
        elif v.endswith("DY"):
            src = v[:-2] + "DTC"
        else:
            return ""
        return (
            f"Equal to {src} - DM.RFSTDTC + 1 if {src} is on or after DM.RFSTDTC;\n"
            f"Equal to {src} - DM.RFSTDTC if {src} precedes DM.RFSTDTC"
        )

    mask = non_crf_mask(var.str.endswith("DY"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = var[mask].apply(build_dy_method)

    # -------------------------------------------------
    # LOBXFL
    # -------------------------------------------------
    mask = non_crf_mask(var.str.endswith("LOBXFL"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = (
        'Equal to "Y" for last record with non-missing value on or before the first exposure date (DM.RFSTDTC);\n'
        'Null otherwise'
    )

    # -------------------------------------------------
    # RDOMAIN 特例
    # -------------------------------------------------
    # CO 特例
    mask = non_crf_mask((ds == "CO") & (var == "RDOMAIN"))
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Codelist"] = "RDOMAIN_CO"

    # SUPP-- 通用規則
    mask = non_crf_mask(ds.str.startswith("SUPP") & (var == "RDOMAIN"))

    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"

    # 把 SUPPxx -> xx
    out.loc[mask, "Codelist"] = ds[mask].str.replace(r"^SUPP", "DOMAIN_", regex=True)


    # -------------------------------------------------
    # CO 保留欄位
    # -------------------------------------------------
    co_assigned_vars = {"RDOMAIN", "IDVAR", "IDVARVAL", "COREF", "COEVAL"}
    mask = non_crf_mask((ds == "CO") & var.isin(co_assigned_vars))
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"

    # -------------------------------------------------
    # DM 特規
    # -------------------------------------------------
    mask = non_crf_mask((ds == "DM") & (var == "RFSTDTC"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = (
        "Equal to date/time of first exposure to study treatment (the earliest value of EXSTDTC)"
    )

    mask = non_crf_mask((ds == "DM") & (var == "RFENDTC"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = (
        "Equal to date/time of last exposure to study treatment (the latest value of EXENDTC)"
    )

    mask = non_crf_mask((ds == "DM") & (var == "RFXSTDTC"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = "Equal to RFSTDTC"

    mask = non_crf_mask((ds == "DM") & (var == "RFXENDTC"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = "Equal to RFENDTC"

    mask = non_crf_mask((ds == "DM") & (var == "RFPENDTC"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = "Equal to the last known date during the study"

    mask = non_crf_mask((ds == "DM") & (var == "DTHFL"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = 'Set to "Y" if DTHDTC is populated'

    mask = non_crf_mask((ds == "DM") & (var == "ARMNRS"))
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"

    mask = non_crf_mask((ds == "DM") & (var == "ACTARMUD"))
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"

    mask = non_crf_mask((ds == "DM") & (var == "COUNTRY"))
    out.loc[mask, "Codelist"] = "ISO3166"

    # -------------------------------------------------
    # DS 特規
    # -------------------------------------------------
    mask = non_crf_mask((ds == "DS") & (var == "DSDTC"))
    out.loc[mask, "Origin"] = "Derived"
    out.loc[mask, "Source"] = "Sponsor"
    out.loc[mask, "Method"] = "Equal to DSSTDTC"

    mask = non_crf_mask((ds == "DS") & (var == "DSCAT"))
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"

    # -------------------------------------------------
    # EC/EX 特規
    # -------------------------------------------------
    EX_vars = {
        "ECTRT", "ECDOSE", "ECDOSU", "ECDOSFRM",
        "EXTRT", "EXDOSE", "EXDOSU", "EXDOSFRM"
    }
    mask = non_crf_mask(var.isin(EX_vars))
    out.loc[mask, "Origin"] = "Protocol"
    out.loc[mask, "Source"] = "Sponsor"


    
    # -------------------------------------------------
    # Protocol-driven trial design vars
    # -------------------------------------------------
    protocol_vars = {
        ("TA", "ELEMENT"), ("TA", "TABRANCH"), ("TA", "TATRANS"),
        ("TE", "ELEMENT"), ("TE", "TESTRL"), ("TE", "TEENRL"), ("TE", "TEDUR"),
        ("TI", "IETEST"), ("TI", "TIVERS"),
        ("TV", "VISITDY"), ("TV", "TVSTRL"), ("TV", "TVENRL"),
    }
    protocol_mask = pd.Series([(d, v) in protocol_vars for d, v in zip(ds, var)], index=out.index)
    protocol_mask = non_crf_mask(protocol_mask)
    out.loc[protocol_mask, "Origin"] = "Protocol"
    out.loc[protocol_mask, "Source"] = "Sponsor"

    # -------------------------------------------------
    # Common Assigned/Sponsor helper vars
    # -------------------------------------------------
    assigned_vars = {
        "IDVAR", "IDVARVAL", "QNAM", "QLABEL", "QORIG", "QEVAL", "COEVAL",
        "ETCD", "TAETORD", "ARMCD", "ARM", "ACTARMCD", "ACTARM",
        "IETESTCD", "EGTESTCD", "EGTEST", "VSTESTCD", "VSTEST",
        "TSPARMCD", "TSPARM", "TSVALCD", "TSVCDREF", "TSVCDVER",
        "AGEU", "DSCAT"
    }
    mask = non_crf_mask(var.isin(assigned_vars))
    out.loc[mask, "Origin"] = "Assigned"
    out.loc[mask, "Source"] = "Sponsor"

    # -------------------------------------------------
    # Common Collected/Investigator vars
    # -------------------------------------------------
    collected_vars = {
        "AESPID", "AETERM", "AESER", "AEACN", "AEREL", "AEOUT", "AESCONG",
        "AESDISAB", "AESDTH", "AESHOSP", "AESLIFE", "AESMIE", "AECONTRT",
        "AETOXGR", "DSTERM", "MHTERM", "RFICDTC", "BRTHDTC", "AGE", "SEX",
        "RACE", "COVAL", "QVAL", "EGORRES", "EGREASND", "EGCLSIG",
        "VSORRES", "VSNRIND", "VSREASND", "VSCLSIG",
        "ECSTDTC", "ECENDTC", "EXSTDTC", "EXENDTC",
        "MHSTDTC", "MHENDTC", "AESTDTC", "AEENDTC",
        "EGDTC", "VSDTC", "EGTPT", "MHSTRTPT", "MHENRTPT",
        "AEENRTPT", "DSSTDTC"
    }
    mask = non_crf_mask(var.isin(collected_vars))
    out.loc[mask, "Origin"] = "Collected"
    out.loc[mask, "Source"] = "Investigator"



    # -------------------------------------------------
    # Comment overrides
    # -------------------------------------------------
    ds = out["Dataset"].astype(str).str.upper()
    var = out["Variable"].astype(str).str.upper()

    # 1) 所有 VISITNUM
    mask = var == "VISITNUM"
    out.loc[mask, "Comment"] = "Assigned from the TV domain based on the VISIT"

    # 2) TA.EPOCH
    mask = (ds == "TA") & (var == "EPOCH")
    out.loc[mask, "Comment"] = "Assigned based on protocol design"

    # 3) IDVAR
    mask = var == "IDVAR"
    out.loc[mask, "Comment"] = (
        "Name of the variables for the related records, such as --SEQ, VISIT or --DTC in related domain"
    )

    # 4) IDVARVAL
    mask = var == "IDVARVAL"
    out.loc[mask, "Comment"] = "Value of identifying variable described in IDVAR"

    

    # =================================================
    # FINAL RULES（一定要放最後）
    # =================================================

    ds = out["Dataset"].astype(str).str.upper()
    var = out["Variable"].astype(str).str.upper()
    origin = out["Origin"].astype(str).str.upper()

    # -------------------------------------------------
    # RULE 1：--TEST / --TESTCD
    # 如果目前 Origin != Collected，則直接 Assigned / Sponsor
    # 例外：TI.IETEST 不套用
    # -------------------------------------------------
    test_mask = (
        (var.str.endswith("TEST") | var.str.endswith("TESTCD")) &
        ~((ds == "TI") & (var == "IETEST"))
    )

    assign_test_mask = test_mask & (origin != "COLLECTED")

    out.loc[assign_test_mask, "Origin"] = "Assigned"
    out.loc[assign_test_mask, "Source"] = "Sponsor"

    # -------------------------------------------------
    # RULE 2：--TEST / --TESTCD 的 Codelist fallback
    # 不看 Origin，只要 codelist 為空就直接等於 variable 本身
    # -------------------------------------------------
    test_mask2 = (
        (var.str.endswith("TEST") | var.str.endswith("TESTCD"))
    )

    empty_cl_mask = test_mask2 & (
        out["Codelist"].fillna("").astype(str).str.strip() == ""
    )
    out.loc[empty_cl_mask, "Codelist"] = var[empty_cl_mask]

    # -------------------------------------------------
    # RULE 3
    # -------------------------------------------------

    # Protocol → Sponsor
    mask = origin == "PROTOCOL"
    out.loc[mask, "Source"] = "Sponsor"

    # Derived → Sponsor
    mask = origin == "DERIVED"
    out.loc[mask, "Source"] = "Sponsor"

    # Assigned → Sponsor（排除 AE dictionary）
    ae_dict_vars = {
        "AELLT", "AELLTCD", "AEDECOD", "AEPTCD",
        "AEHLT", "AEHLTCD", "AEHLGT", "AEHLGTCD",
        "AEBODSYS", "AEBDSYCD", "AESOC", "AESOCCD"
    }
    
    mask = (origin == "ASSIGNED") & (~var.isin(ae_dict_vars))
    out.loc[mask, "Source"] = "Sponsor"


    
    return out



def build_variables_spec_from_domains_config(detail_df, config_df):
    detected_datasets = (
        sorted(detail_df["SDTM Domain"].astype(str).str.upper().unique())
        if not detail_df.empty else []
    )
    expanded_cfg = expand_suppqual_to_supp_datasets(config_df, detected_datasets)

    # ---------------------------
    # Part 1: CRF collected / assigned variables
    # ---------------------------
    crf_part = enrich_crf_variables_with_config(detail_df, expanded_cfg)

    # ---------------------------
    # Part 2: non-CRF variables from config
    # ---------------------------
    non_crf_part = get_non_crf_from_config(detail_df, expanded_cfg)

    final_df = pd.concat([crf_part, non_crf_part], ignore_index=True)
    final_df = final_df.drop_duplicates(subset=["Dataset", "Variable"], keep="first")

    # ---------------------------
    # Part 3: Trial Design variables（跟 config 比對）
    # ---------------------------
    td_var_df = build_trial_design_variables_spec(expanded_cfg)
    final_df = pd.concat([final_df, td_var_df], ignore_index=True)
    final_df = final_df.drop_duplicates(subset=["Dataset", "Variable"], keep="first")

    # ---------------------------
    # Part 4: append required partner variables
    # 只加 config 裡存在的變數
    # ---------------------------
    final_df = append_required_partner_variables(final_df, expanded_cfg)

    # ---------------------------
    # Part 5: type / base codelist overrides
    # ---------------------------
    final_df = apply_variable_level_overrides(final_df)

    # ---------------------------
    # Part 6: final Origin / Source / Method / special codelist overrides
    # ---------------------------
    final_df = apply_origin_source_method_overrides(final_df)

    # ---------------------------
    # Part 7: sort by Dataset + VarNum + Variable
    # Order 每個 Dataset 重新從 1 開始
    # ---------------------------
    if "VarNum" not in final_df.columns:
        final_df["VarNum"] = ""

    final_df["VarNum_num"] = pd.to_numeric(final_df["VarNum"], errors="coerce")

    final_df = final_df.sort_values(
        by=["Dataset", "VarNum_num", "Variable"],
        na_position="last"
    ).reset_index(drop=True)

    final_df["Order"] = final_df.groupby("Dataset").cumcount() + 1

    keep_cols = [
        "Order", "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment"
    ]

    for c in keep_cols:
        if c not in final_df.columns:
            final_df[c] = ""

    final_df = final_df[keep_cols]

    return final_df



def apply_variable_level_overrides(df):
    if df.empty:
        return df

    out = df.copy()

    for c in ["Dataset", "Variable", "Codelist", "Data Type"]:
        if c not in out.columns:
            out[c] = ""

    ds_upper = out["Dataset"].astype(str).str.upper()
    var_upper = out["Variable"].astype(str).str.upper()

    # 先不要太早固定 cl_upper，因為後面 Codelist 會被改寫
    # cl_upper 如有需要，務必在後面重新計算

    # ---------------------------
    # 1) Data Type normalize
    # ---------------------------
    out["Data Type"] = out.apply(
        lambda r: normalize_data_type_by_config(r.get("Data Type", ""), r.get("Variable", "")),
        axis=1
    )

    # ---------------------------
    # 2) variable-driven special rules
    #    這些不能只靠原始 Codelist 判斷
    # ---------------------------
    # DOMAIN
    mask = var_upper == "DOMAIN"
    out.loc[mask, "Codelist"] = ds_upper[mask].apply(lambda x: f"DOMAIN_{x}")

    # STRTPT / ENRTPT
    mask = var_upper.str.endswith("STRTPT")
    out.loc[mask, "Codelist"] = ds_upper[mask].apply(lambda x: f"STENRF_{x}_START")

    mask = var_upper.str.endswith("ENRTPT")
    out.loc[mask, "Codelist"] = ds_upper[mask].apply(lambda x: f"STENRF_{x}_END")

    # EPOCH
    mask = var_upper == "EPOCH"
    out.loc[mask, "Codelist"] = "EPOCH"

    # AE dictionary
    ae_dict_vars = {
        "AELLT", "AELLTCD", "AEDECOD", "AEPTCD",
        "AEHLT", "AEHLTCD", "AEHLGT", "AEHLGTCD",
        "AEBODSYS", "AEBDSYCD", "AESOC", "AESOCCD"
    }
    out.loc[var_upper.isin(ae_dict_vars), "Codelist"] = "AEDICT_F"

    # AEREL
    out.loc[var_upper == "AEREL", "Codelist"] = "AEREL"

    # DSDECOD
    out.loc[var_upper == "DSDECOD", "Codelist"] = "NCOMPLT"


    var_upper = out["Variable"].astype(str).str.upper()
    
    # ARMCD / ACTARMCD → ARMCD
    mask = var_upper.isin(["ARMCD", "ACTARMCD"])
    out.loc[mask, "Codelist"] = "ARMCD"

    # ARM / ACTARM → ARM
    mask = var_upper.isin(["ARM", "ACTARM"])
    out.loc[mask, "Codelist"] = "ARM"

    
    # ---------------------------
    # 3) base codelist behavior table
    # ---------------------------
    CODELIST_BEHAVIOR = {
        "UNIT": "SUFFIX_DOMAIN",
        "FRM": "SUFFIX_DOMAIN",
        "ARMCD": "KEEP",
        "ARM": "KEEP",
        "NY": "KEEP",
        "Y": "KEEP",
        "ND": "KEEP",
        "EPOCH": "KEEP",
        "IETESTCD": "KEEP",
        "IETEST": "KEEP",
        "TSPARMCD": "KEEP",
        "TSPARM": "KEEP",
    }

    # 這裡再重新抓一次，因為前面已經改過 Codelist
    cl_upper = out["Codelist"].astype(str).str.strip().str.upper()

    for i in out.index:
        code = cl_upper[i]
        ds = ds_upper[i]

        if code in ["", "NAN", "NONE"]:
            continue

        behavior = CODELIST_BEHAVIOR.get(code, "KEEP")

        if behavior == "SUFFIX_DOMAIN":
            out.at[i, "Codelist"] = f"{code}_{ds}"

    return out



def make_empty_variable_row(dataset, variable):
    return {
        "Dataset": dataset,
        "Variable": variable,
        "Label": "",
        "Data Type": "",
        "Codelist": "",
        "Origin": "Derived",
        "Source": "",
        "Pages": "",
        "Method": "",
        "Comment": "",
        "Core": "",
        "VarNum": ""
    }



def build_datasets_spec_from_domains_config(mapping_df, config_df, version):
    detected_datasets = []

    if not mapping_df.empty:
        detected_datasets = sorted(mapping_df["SDTM Domain"].dropna().astype(str).str.upper().unique())

    expanded_cfg = expand_suppqual_to_supp_datasets(config_df, detected_datasets)

    ds_cols = [c for c in [
        "Dataset", "Dataset Label", "Class", "Structure", "Key Variables"
    ] if c in expanded_cfg.columns]

    if ds_cols:
        ds_df = expanded_cfg[ds_cols].drop_duplicates(subset=["Dataset"]).copy()
        if detected_datasets:
            ds_df = ds_df[ds_df["Dataset"].isin(detected_datasets)]
    else:
        ds_df = pd.DataFrame(columns=["Dataset", "Dataset Label", "Class", "Structure", "Key Variables"])

    ds_df = ds_df.rename(columns={
        "Dataset Label": "Label"
    })

    if "Label" not in ds_df.columns:
        ds_df["Label"] = ""
    if "Class" not in ds_df.columns:
        ds_df["Class"] = ""
    if "Structure" not in ds_df.columns:
        ds_df["Structure"] = ""
    if "Key Variables" not in ds_df.columns:
        ds_df["Key Variables"] = ""

    std_ver = version.replace("Version", "").strip()
    ds_df["Standard"] = f"SDTMIG {std_ver}"

    ds_df = ds_df[[
        "Dataset", "Label", "Class", "Structure", "Key Variables", "Standard"
    ]].reset_index(drop=True)

    # ---------------------------
    # Append Trial Design datasets
    # ---------------------------
    td_df = build_trial_design_datasets_spec(version)

    final_df = pd.concat([ds_df, td_df], ignore_index=True)
    final_df = final_df.drop_duplicates(subset=["Dataset"], keep="first")
    final_df = final_df.sort_values(by=["Dataset"]).reset_index(drop=True)

    return final_df



# =========================================================
# Define / Codelists / Dictionaries / Trial Design
# =========================================================
def extract_protocol_no_from_filename(file_name):
    if not file_name:
        return ""

    name = os.path.splitext(file_name)[0].strip()

    parts = re.split(r"ecrf\s*schema", name, flags=re.IGNORECASE)
    prefix = parts[0].strip().strip("_") if parts else name

    tokens = [t.strip() for t in prefix.split("_") if t.strip()]

    if not tokens:
        return ""

    protocol = tokens[-1]
    return protocol


def build_define_sheet(version, protocol_no="", protocol_title=""):
    std_ver = version.replace("Version", "").strip()

    define_df = pd.DataFrame({
        "Attribute": [
            "StudyName",
            "StudyDescription",
            "ProtocolName",
            "StandardName",
            "StandardVersion",
            "Language"
        ],
        "Value": [
            protocol_no,
            protocol_title,
            protocol_no,
            "SDTM-IG",
            std_ver,
            "en"
        ]
    })

    return define_df


def build_codelists_sheet_from_variables(variables_df):
    cols = [
        "ID", "Name", "NCI Codelist Code", "Data Type", "Terminology",
        "Comment", "Order", "Term", "NCI Term Code", "Decoded Value"
    ]

    if variables_df.empty or "Codelist" not in variables_df.columns:
        return pd.DataFrame(columns=cols)

    ids = (
        variables_df["Codelist"]
        .dropna()
        .astype(str)
        .str.strip()
    )
    ids = ids[ (ids != "") & (ids.str.upper() != "AEDICT_F")]
    ids = sorted(ids.drop_duplicates().tolist())

    return pd.DataFrame({
        "ID": ids,
        "Name": ["" for _ in ids],
        "NCI Codelist Code": ["" for _ in ids],
        "Data Type": ["" for _ in ids],
        "Terminology": ["" for _ in ids],
        "Comment": ["" for _ in ids],
        "Order": ["" for _ in ids],
        "Term": ["" for _ in ids],
        "NCI Term Code": ["" for _ in ids],
        "Decoded Value": ["" for _ in ids]
    })


def build_default_dictionaries_sheet(meddra_version="", cm_dictionary="WHO ATC/DDD", cm_version=""):
    return pd.DataFrame([
        {
            "ID": "AEDICT_F",
            "Name": "Adverse Event Dictionary",
            "Data Type": "text",
            "Dictionary": "MEDDRA",
            "Version": meddra_version
        },
        {
            "ID": "CMDICT_F",
            "Name": "Concomitant Meds Dictionary",
            "Data Type": "text",
            "Dictionary": cm_dictionary,
            "Version": cm_version
        },
        {
            "ID": "ISO3166",
            "Name": "Country Codes (ISO 3166)",
            "Data Type": "text",
            "Dictionary": "ISO 3166",
            "Version": ""
        }
    ])



def get_trial_design_definitions():
    return {
        "TA": {
            "label": "Trial Arms",
            "class": "Trial Design",
            "structure": "One record per planned arm",
            "key_variables": "STUDYID, ARMCD",
            "variables": [
                ("STUDYID", "Study Identifier", "text"),
                ("DOMAIN", "Domain Abbreviation", "text"),
                ("ARMCD", "Planned Arm Code", "text"),
                ("ARM", "Description of Planned Arm", "text"),
                ("TAETORD", "Planned Order of Elements Within Arm", "integer"),
                ("ETCD", "Element Code", "text"),
                ("ELEMENT", "Description of Element", "text"),
                ("TABRANCH", "Branch", "text"),
                ("TATRANS", "Transition Rule", "text"),
                ("EPOCH", "Epoch", "text"),            
            ]
        },
        "TE": {
            "label": "Trial Elements",
            "class": "Trial Design",
            "structure": "One record per element",
            "key_variables": "STUDYID, ETCD",
            "variables": [
                ("STUDYID", "Study Identifier", "text"),
                ("DOMAIN", "Domain Abbreviation", "text"),
                ("ETCD", "Element Code", "text"),
                ("ELEMENT", "Description of Element", "text"),
                ("TESTRL", "Rule for Start of Element", "text"),
                ("TEENRL", "End Rule", "text"),
                ("TEDUR", "Planned Duration of Element", "text"),
            ]
        },
        "TI": {
            "label": "Trial Inclusion/Exclusion Criteria",
            "class": "Trial Design",
            "structure": "One record per inclusion/exclusion criterion",
            "key_variables": "STUDYID, IETESTCD",
            "variables": [
                ("STUDYID", "Study Identifier", "text"),
                ("DOMAIN", "Domain Abbreviation", "text"),
                ("IETESTCD", "Inclusion/Exclusion Criterion Short Name", "text"),
                ("IETEST", "Inclusion/Exclusion Criterion", "text"),
                ("IECAT", "Inclusion/Exclusion Category", "text"),
                ("TIVERS", "Version", "text"), 
            ]
        },
        "TS": {
            "label": "Trial Summary",
            "class": "Trial Design",
            "structure": "One record per trial summary parameter",
            "key_variables": "STUDYID, TSSEQ",
            "variables": [
                ("STUDYID", "Study Identifier", "text"),
                ("DOMAIN", "Domain Abbreviation", "text"),
                ("TSSEQ", "Sequence Number", "integer"),
                ("TSPARMCD", "Trial Summary Parameter Short Name", "text"),
                ("TSPARM", "Trial Summary Parameter", "text"),
                ("TSVAL", "Parameter Value", "text"),
                ("TSVALCD", "Parameter Value (Code)", "text"),
                ("TSVCDREF", "Code Dictionary Reference", "text"),
                ("TSVCDVER", "Code Dictionary Version", "text"), 
                ("TSVALNF", "Null Flavor", "text"),
            ]
        },
        "TV": {
            "label": "Trial Visits",
            "class": "Trial Design",
            "structure": "One record per visit per arm",
            "key_variables": "STUDYID, VISITNUM",
            "variables": [
                ("STUDYID", "Study Identifier", "text"),
                ("DOMAIN", "Domain Abbreviation", "text"),
                ("VISITNUM", "Visit Number", "float"),
                ("VISIT", "Visit Name", "text"),
                ("VISITDY", "Planned Study Day of Visit", "integer"),
                ("ARMCD", "Planned Arm Code", "text"),
                ("ARM", "Planned Arm", "text"), 
                ("TVSTRL", "Start Rule", "text"), 
                ("TVENRL", "End Rule", "text"),

            ]
        }
    }



def get_trial_design_columns_from_config(domain, config_df, fallback_columns):
    """
    依 config 中的 VarNum 排 Trial Design template 欄位順序。
    若 config 沒有某些欄位，仍保留並放在最後（依 fallback 原順序）。
    """
    domain = str(domain).upper()
    fallback_columns = [str(c).upper() for c in fallback_columns]

    if config_df is None or config_df.empty:
        return fallback_columns

    cfg = config_df.copy()

    required_cols = {"Dataset", "Variable"}
    if not required_cols.issubset(set(cfg.columns)):
        return fallback_columns

    td_cfg = cfg[cfg["Dataset"].astype(str).str.upper() == domain].copy()

    if td_cfg.empty:
        return fallback_columns

    if "VarNum" in td_cfg.columns:
        td_cfg["VarNum_num"] = pd.to_numeric(td_cfg["VarNum"], errors="coerce")
    else:
        td_cfg["VarNum_num"] = pd.NA

    td_cfg["Variable"] = td_cfg["Variable"].astype(str).str.upper()

    # 只取 fallback_columns 裡有定義的欄位，避免 config 多出不想顯示的內容
    td_cfg = td_cfg[td_cfg["Variable"].isin(fallback_columns)].copy()

    td_cfg = td_cfg.sort_values(
        by=["VarNum_num", "Variable"],
        na_position="last"
    )

    ordered_from_cfg = td_cfg["Variable"].dropna().astype(str).str.upper().tolist()

    # 把 config 沒列到的欄位補到後面，維持 fallback 順序
    remaining = [c for c in fallback_columns if c not in ordered_from_cfg]

    final_cols = ordered_from_cfg + remaining

    # 去重，保留順序
    final_cols = list(dict.fromkeys(final_cols))

    return final_cols


def build_trial_design_templates(protocol_no="", config_df=None):
    """
    Trial Design template:
      - STUDYID 自動帶 protocol_no
      - DOMAIN 自動帶 TA/TE/TI/TS/TV
      - 欄位順序依 config 的 VarNum 排
    """
    defs = get_trial_design_definitions()
    outputs = []

    for domain in ["TA", "TE", "TI", "TS", "TV"]:
        fallback_columns = [v[0].upper() for v in defs[domain]["variables"]]
        ordered_columns = get_trial_design_columns_from_config(
            domain=domain,
            config_df=config_df,
            fallback_columns=fallback_columns
        )

        row = {c: "" for c in ordered_columns}

        if "STUDYID" in row:
            row["STUDYID"] = protocol_no
        if "DOMAIN" in row:
            row["DOMAIN"] = domain

        df = pd.DataFrame([row], columns=ordered_columns)
        outputs.append(df)

    return tuple(outputs)



def build_trial_design_datasets_spec(version):
    defs = get_trial_design_definitions()
    std_ver = version.replace("Version", "").strip()

    rows = []
    for domain in ["TA", "TE", "TI", "TS", "TV"]:
        info = defs[domain]
        rows.append({
            "Dataset": domain,
            "Label": info["label"],
            "Class": info["class"],
            "Structure": info["structure"],
            "Key Variables": info["key_variables"],
            "Standard": f"SDTMIG {std_ver}"
        })

    return pd.DataFrame(rows)




def to_excel_bytes(sheet_dict):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheet_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output.getvalue()


# =========================================================
# 共用：從 Excel 建立 CRF mapping
# =========================================================
def process_uploaded_excel(file_bytes, all_sheets, manual_soa_header=None, common_domain_header=None):
    if "SoA" not in all_sheets:
        raise ValueError("找不到 SoA 分頁")

    soa_df, _ = read_sheet_with_detected_header(
        file_bytes=file_bytes,
        sheet_name="SoA",
        keyword_groups=[["FORM", "OID"]],
        manual_header_row_excel=manual_soa_header
    )

    form_oid_col = find_column(soa_df.columns, ["FORM", "OID"])
    if form_oid_col is None:
        raise ValueError("SoA 分頁中找不到 Form OID 欄位")

    valid_domains = extract_form_oids(soa_df[form_oid_col])
    sheet_upper_map = {s.upper(): s for s in all_sheets}

    available_sheets = [
        sheet_upper_map[d] for d in valid_domains if d in sheet_upper_map
    ]

    missing_sheets = [
        d for d in valid_domains if d not in sheet_upper_map
    ]

    mapping_df, detail_df, sheet_errors, unparsed_records = build_sdtm_mapping(
        file_bytes=file_bytes,
        selected_crf_sheets=available_sheets,
        common_domain_header=common_domain_header
    )
    
    ct_mapping_df, ct_mapping_sheet_errors = build_ct_mapping_seed(
        file_bytes=file_bytes,
        selected_crf_sheets=available_sheets,
        common_domain_header=common_domain_header
    )

    return {
        "available_sheets": available_sheets,
        "missing_sheets": missing_sheets,
        "mapping_df": mapping_df,
        "detail_df": detail_df,
        "sheet_errors": sheet_errors,
        "unparsed_records": unparsed_records,
        "ct_mapping_df": ct_mapping_df,
        "ct_mapping_sheet_errors": ct_mapping_sheet_errors
    }


# =========================================================
# Step 1 cache key
# =========================================================
def make_step1_cache_key(file_bytes, manual_soa_header, common_domain_header):
    md5 = hashlib.md5(file_bytes).hexdigest()
    return f"{md5}|soa={manual_soa_header}|domain={common_domain_header}"


# =========================================================
# 主流程 UI
# =========================================================
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

        step1_cache_key = make_step1_cache_key(
            file_bytes=file_bytes,
            manual_soa_header=manual_soa_header,
            common_domain_header=common_domain_header
        )

        if (
            st.session_state.get("step1_cache_key") == step1_cache_key
            and st.session_state.get("step1_result") is not None
        ):
            result = st.session_state["step1_result"]
        else:
            result = process_uploaded_excel(
                file_bytes=file_bytes,
                all_sheets=all_sheets,
                manual_soa_header=manual_soa_header,
                common_domain_header=common_domain_header
            )
            st.session_state["step1_cache_key"] = step1_cache_key
            st.session_state["step1_result"] = result

        missing_sheets = result["missing_sheets"]
        mapping_df = result["mapping_df"]
        detail_df = result["detail_df"]
        sheet_errors = result["sheet_errors"]
        unparsed_records = result["unparsed_records"]
        ct_mapping_df = result.get("ct_mapping_df", pd.DataFrame())
        ct_mapping_sheet_errors = result.get("ct_mapping_sheet_errors", [])


        # 給 Step 2 用
        st.session_state["mapping_df"] = mapping_df
        st.session_state["detail_df"] = detail_df

        if missing_sheets:
            st.warning(f"SoA 有但 Excel 沒有的 Sheets：{missing_sheets}")

        st.markdown("### 整份檔案要呈現的 SDTM Domains / Variables")
        if mapping_df.empty:
            st.warning("目前沒有從各 CRF sheet 的 SDTM IG Target 抓到可解析的 SDTM domain / variable")
        else:
            summary_df = summarize_sdtm_mapping(mapping_df)
            st.dataframe(summary_df, use_container_width=True)

        st.markdown("### SDTM Mapping 明細")
        if detail_df.empty:
            st.info("目前沒有可顯示的明細")
        else:
            st.dataframe(detail_df, use_container_width=True)

        if sheet_errors:
            clean_sheets = sorted(set(sheet_errors))
            st.markdown("### 無法處理的 Sheets")
            st.warning(f"header 偵測失敗，無法自動判斷 header row: {clean_sheets}")

        if unparsed_records:
            st.markdown("### 無法解析的 SDTM IG Target 值")
            st.dataframe(pd.DataFrame(unparsed_records), use_container_width=True)



       st.markdown("### CT Term Mapping List（from CRF Options）")
        if ct_mapping_df.empty:
            st.info("目前尚未從 CRF schema 抽到 Option Displayed Value")
        else:
            ct_mapping_df = st.data_editor(
                ct_mapping_df,
                num_rows="dynamic",
                use_container_width=True,
                key="ct_mapping_editor"
            )

        # 存給 Step 2 用
        st.session_state["ct_mapping_df"] = ct_mapping_df

        if ct_mapping_sheet_errors:
            st.caption(f"以下 sheets 無法抽取 CT option（可能沒有 option 欄位或 header 偵測失敗）：{ct_mapping_sheet_errors}")


        # -------------------------------------------------
        # Step 2 開關：執行 / 重新整理
        # -------------------------------------------------
        def trigger_step2():
            st.session_state["run_step2"] = True

        st.button(
            "▶ 執行 / 重新整理 Step 2：SPEC Generator",
            type="primary",
            on_click=trigger_step2
        )

        # -------------------------------------------------
        # Step 2：SPEC Generator
        # -------------------------------------------------
        if st.session_state.get("run_step2", False):
            st.markdown("## Step 2｜SPEC Generator")

            mapping_df = st.session_state.get("mapping_df", pd.DataFrame())
            detail_df = st.session_state.get("detail_df", pd.DataFrame())

            if mapping_df.empty:
                st.warning("目前沒有可用的 CRF → SDTM mapping，無法建立 SPEC")
            else:
                st.markdown("### Basic Information")

                default_protocol_no = extract_protocol_no_from_filename(uploaded_file.name)

                col1, col2 = st.columns(2)
                with col1:
                    protocol_no = st.text_input(
                        "Protocol No",
                        value=default_protocol_no,
                        key="protocol_no"
                    )
                with col2:
                    protocol_title = st.text_input(
                        "Protocol Title (請填寫)",
                        value="",
                        key="protocol_title"
                    )

                st.markdown("##### Version Control")

                # ---------------------------
                # Row 1：SDTM IG + SDTM CT
                # ---------------------------
                r1_c1, r1_c2 = st.columns(2)

                with r1_c1:
                    version = st.selectbox(
                        "SDTM IG",
                        ["Version 3.4", "Version 3.3"],
                        key="sdtm_version_selector"
                    )

                with r1_c2:
                    sdtm_ct = st.text_input(
                        "SDTM CT",
                        value="",
                        key="sdtm_ct"
                    )

                # ---------------------------
                # Row 2：MedDRA（單獨一列）
                # ---------------------------
                meddra_version = st.text_input(
                    "MedDRA",
                    value="",
                    key="meddra_version"
                )

                # ---------------------------
                # Row 3：CM 字典 + CM 版本
                # ---------------------------
                r3_c1, r3_c2 = st.columns(2)

                with r3_c1:
                    cm_dictionary = st.selectbox(
                        "CM 字典",
                        ["WHODrug Global B3", "WHO ATC/DDD"],
                        key="cm_dictionary"
                    )

                with r3_c2:
                    cm_version = st.text_input(
                        "CM 版本",
                        value="",
                        key="cm_version"
                    )

                # ---------------------------
                # Row 4：其他（SNOMED / UNII / MED-RT）
                # ---------------------------
                r4_c1, r4_c2, r4_c3 = st.columns(3)

                with r4_c1:
                    snomed_version = st.text_input(
                        "SNOMED",
                        value="",
                        key="snomed_version"
                    )

                with r4_c2:
                    unii_version = st.text_input(
                        "UNII",
                        value="",
                        key="unii_version"
                    )

                with r4_c3:
                    medrt_version = st.text_input(
                        "MED-RT",
                        value="",
                        key="medrt_version"
                    )


                try:
                    raw_cfg_df, cfg_path = load_domains_config(version)
                    cfg_df = standardize_domains_config(raw_cfg_df)

                    # st.success(f"✅ 已成功載入 config：{cfg_path}")

                    # 2.1 Define
                    st.markdown("### 2.1 Define")
                    define_df = build_define_sheet(
                        version=version,
                        protocol_no=protocol_no,
                        protocol_title=protocol_title
                    )
                    st.dataframe(define_df, use_container_width=True)

                    # 2.2 Datasets
                    st.markdown("### 2.2 Datasets")
                    datasets_spec_df = build_datasets_spec_from_domains_config(
                        mapping_df=mapping_df,
                        config_df=cfg_df,
                        version=version
                    )
                    st.dataframe(datasets_spec_df, use_container_width=True)

                    # 2.3 Variables
                    st.markdown("### 2.3 Variables")
                    variables_spec_df = build_variables_spec_from_domains_config(
                        detail_df=detail_df,
                        config_df=cfg_df
                    )
                    st.dataframe(variables_spec_df, use_container_width=True)

                    # 2.4 Codelists
                    st.markdown("### 2.4 Codelists")
                    codelists_df = build_codelists_sheet_from_variables(variables_spec_df)
                    st.dataframe(codelists_df, use_container_width=True)

                    # 2.5 Dictionaries
                    st.markdown("### 2.5 Dictionaries")
                    dictionaries_df = st.data_editor(
                        build_default_dictionaries_sheet(
                            meddra_version=meddra_version,
                            cm_dictionary=cm_dictionary,
                            cm_version=cm_version
                        ),
                        num_rows="dynamic",
                        use_container_width=True,
                        key="dictionaries_editor"
                    )

                    # 2.6 Trial Design
                    st.markdown("### 2.6 Trial Design")
                    ta_df, te_df, ti_df, ts_df, tv_df = build_trial_design_templates(
                        protocol_no=protocol_no,
                        config_df=cfg_df
                    )

                    with st.expander("TA / TE / TI / TS / TV 基本欄位骨架", expanded=False):
                        st.markdown("#### TA")
                        st.dataframe(ta_df, use_container_width=True)

                        st.markdown("#### TE")
                        st.dataframe(te_df, use_container_width=True)

                        st.markdown("#### TI")
                        st.dataframe(ti_df, use_container_width=True)

                        st.markdown("#### TS")
                        st.dataframe(ts_df, use_container_width=True)

                        st.markdown("#### TV")
                        st.dataframe(tv_df, use_container_width=True)

                    export_sheets = {
                        "Define": define_df,
                        "Datasets": datasets_spec_df,
                        "Variables": variables_spec_df,
                        "Codelists": codelists_df,
                        "Dictionaries": dictionaries_df,
                        "TA": ta_df,
                        "TE": te_df,
                        "TI": ti_df,
                        "TS": ts_df,
                        "TV": tv_df
                    }

                    excel_bytes = to_excel_bytes(export_sheets)

                    st.download_button(
                        label="下載 SDTM SPEC Excel",
                        data=excel_bytes,
                        file_name=f"SDTM_SPEC_{version.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"Step 2 載入 config / 產生 SPEC 失敗：{e}")

    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
