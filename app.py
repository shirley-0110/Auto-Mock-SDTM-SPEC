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

    out = non_crf[[
        "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment", "Core", "VarNum"
    ]].drop_duplicates(subset=["Dataset", "Variable"])

    return out.reset_index(drop=True)


def enrich_crf_variables_with_config(detail_df, config_df):
    """
    CRF 收集到的都保留
    並排除 SEND Only
    同一個 Dataset+Variable 只保留一筆
    """
    if detail_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment", "Core", "VarNum"
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
        return "CRF"

    crf_df["Origin"] = crf_df.apply(derive_origin, axis=1)
    crf_df["Source"] = crf_df.apply(
        lambda r: f"{r['Source CRF Sheet']} / {r['Source CRF Variable']}".strip(" /"),
        axis=1
    )
    crf_df["Pages"] = ""
    crf_df["Method"] = ""
    crf_df["Comment"] = crf_df.apply(
        lambda r: f"Assign Value={r['Assign Value']}" if str(r.get("Assign Value", "")).strip() != "" else "",
        axis=1
    )

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
        "Origin": lambda s: "Assigned" if "Assigned" in list(s) else "CRF",
        "Source": join_unique,
        "Pages": "first",
        "Method": "first",
        "Comment": join_unique,
        "Core": "first",
        "VarNum": "first"
    }).reset_index()

    return grouped[[
        "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment", "Core", "VarNum"
    ]]


def build_variables_spec_from_domains_config(detail_df, config_df):
    detected_datasets = sorted(detail_df["SDTM Domain"].astype(str).str.upper().unique()) if not detail_df.empty else []
    expanded_cfg = expand_suppqual_to_supp_datasets(config_df, detected_datasets)

    crf_part = enrich_crf_variables_with_config(detail_df, expanded_cfg)
    non_crf_part = get_non_crf_from_config(detail_df, expanded_cfg)

    final_df = pd.concat([crf_part, non_crf_part], ignore_index=True)
    final_df = final_df.drop_duplicates(subset=["Dataset", "Variable"])

    if "VarNum" in final_df.columns:
        final_df["VarNum_num"] = pd.to_numeric(final_df["VarNum"], errors="coerce")
        final_df = final_df.sort_values(by=["Dataset", "VarNum_num", "Variable"], na_position="last")
        final_df = final_df.drop(columns=["VarNum_num"])
    else:
        final_df = final_df.sort_values(by=["Dataset", "Variable"])

    final_df = final_df.reset_index(drop=True)
    final_df.insert(0, "Order", range(1, len(final_df) + 1))

    final_df = final_df[[
        "Order", "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment", "Core"
    ]]

    return final_df



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
    ids = ids[ids != ""]
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
                ("TSGRPID", "Group ID", "text"),
                ("TSPARMCD", "Trial Summary Parameter Short Name", "text"),
                ("TSPARM", "Trial Summary Parameter", "text"),
                ("TSVAL", "Parameter Value", "text"),
                ("TSVALCD", "Parameter Value (Code)", "text"),
                ("TSVCDREF", "Code Dictionary Reference", "text"),
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
            ]
        }
    }


def build_trial_design_templates(protocol_no=""):
    defs = get_trial_design_definitions()

    outputs = []

    for domain in ["TA", "TE", "TI", "TS", "TV"]:
        cols = [v[0] for v in defs[domain]["variables"]]

        row = {c: "" for c in cols}
        if "STUDYID" in row:
            row["STUDYID"] = protocol_no
        if "DOMAIN" in row:
            row["DOMAIN"] = domain

        df = pd.DataFrame([row], columns=cols)
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


def build_trial_design_variables_spec():
    defs = get_trial_design_definitions()

    rows = []
    for domain in ["TA", "TE", "TI", "TS", "TV"]:
        for var, label, dtype in defs[domain]["variables"]:
            rows.append({
                "Dataset": domain,
                "Variable": var,
                "Label": label,
                "Data Type": dtype,
                "Codelist": "",
                "Origin": "Protocol",
                "Source": "",
                "Pages": "",
                "Method": "",
                "Comment": "Trial Design template",
                "Core": ""
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

    return {
        "available_sheets": available_sheets,
        "missing_sheets": missing_sheets,
        "mapping_df": mapping_df,
        "detail_df": detail_df,
        "sheet_errors": sheet_errors,
        "unparsed_records": unparsed_records
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
                        ["Version 3.3", "Version 3.4"],
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
                    value="28.1",
                    key="meddra_version"
                )

                # ---------------------------
                # Row 3：CM 字典 + CM 版本
                # ---------------------------
                r3_c1, r3_c2 = st.columns(2)

                with r3_c1:
                    cm_dictionary = st.selectbox(
                        "CM 字典",
                        ["WHO ATC/DDD", "WHODrug Global B3"],
                        key="cm_dictionary"
                    )

                with r3_c2:
                    cm_version = st.text_input(
                        "CM 版本",
                        value="2025",
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
                    ta_df, te_df, ti_df, ts_df, tv_df = build_trial_design_templates()

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
