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
    """
    只讀 repo 中的單一 domains.sas7bdat
    路徑：
      config/v33/domains.sas7bdat
      config/v34/domains.sas7bdat
    """
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
    """
    把 domains.sas7bdat 標準化成 app 內部用欄位
    原始欄位預期：
      domain, dlabel, repeat, refdata, structure, keyvars, keyseq,
      varnum, name, label, type, mandatory, role, core, ctcode, class
    """
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
    """
    將 config 中的 SUPPQUAL 展開到所有偵測到的 SUPP-- datasets
    例如：
      SUPPQUAL -> SUPPAE, SUPPDM, SUPPVS ...
    Dataset Label 改成：
      Supplemental Qualifiers for AE
      Supplemental Qualifiers for DM
    """
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


def get_non_crf_from_config(detail_df, config_df):
    """
    non-CRF = config 裡所有 variables - Step 1 mapping 已有的 variables
    只保留目前 mapping 有出現的 datasets
    """
    if detail_df.empty or config_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment",
            "Mandatory", "Role", "Core", "Class", "VarNum"
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

    cfg["pair"] = list(zip(cfg["Dataset"], cfg["Variable"]))
    non_crf = cfg[~cfg["pair"].isin(crf_pairs)].copy()

    if "pair" in non_crf.columns:
        non_crf = non_crf.drop(columns=["pair"])

    non_crf["Label"] = non_crf.get("Variable Label", "")
    non_crf["Origin"] = non_crf.get("Origin", "")
    non_crf["Source"] = non_crf.get("Source", "")
    non_crf["Pages"] = ""
    non_crf["Method"] = non_crf.get("Method", "")
    non_crf["Comment"] = non_crf.get("Comment", "")

    for col in ["Label", "Data Type", "Codelist", "Origin", "Source", "Pages", "Method", "Comment",
                "Mandatory", "Role", "Core", "Class", "VarNum"]:
        if col not in non_crf.columns:
            non_crf[col] = ""

    return non_crf[[
        "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment",
        "Mandatory", "Role", "Core", "Class", "VarNum"
    ]].drop_duplicates().reset_index(drop=True)


def enrich_crf_variables_with_config(detail_df, config_df):
    """
    用 domains.sas7bdat 補 Step 1 抓到的 CRF variables metadata
    SUPP-- 會自動連到 SUPPQUAL 的展開結果
    """
    if detail_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment",
            "Mandatory", "Role", "Core", "Class", "VarNum"
        ])

    detected_datasets = sorted(detail_df["SDTM Domain"].astype(str).str.upper().unique())
    expanded_cfg = expand_suppqual_to_supp_datasets(config_df, detected_datasets)

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
            "Mandatory", "Role", "Core", "Class", "VarNum"
        ] if c in expanded_cfg.columns
    ]

    cfg = expanded_cfg[cfg_keep_cols].drop_duplicates() if cfg_keep_cols else pd.DataFrame(columns=["Dataset", "Variable"])

    merged = crf_df.merge(
        cfg,
        on=["Dataset", "Variable"],
        how="left"
    )

    merged["Label"] = merged.get("Variable Label", "")

    for col in ["Label", "Data Type", "Codelist", "Mandatory", "Role", "Core", "Class", "VarNum"]:
        if col not in merged.columns:
            merged[col] = ""

    return merged[[
        "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment",
        "Mandatory", "Role", "Core", "Class", "VarNum"
    ]].drop_duplicates().reset_index(drop=True)


def build_variables_spec_from_domains_config(detail_df, config_df):
    detected_datasets = sorted(detail_df["SDTM Domain"].astype(str).str.upper().unique()) if not detail_df.empty else []
    expanded_cfg = expand_suppqual_to_supp_datasets(config_df, detected_datasets)

    crf_part = enrich_crf_variables_with_config(detail_df, expanded_cfg)
    non_crf_part = get_non_crf_from_config(detail_df, expanded_cfg)

    final_df = pd.concat([crf_part, non_crf_part], ignore_index=True)
    final_df = final_df.drop_duplicates()

    if "VarNum" in final_df.columns:
        final_df["VarNum_num"] = pd.to_numeric(final_df["VarNum"], errors="coerce")
        final_df = final_df.sort_values(by=["Dataset", "VarNum_num", "Variable"], na_position="last")
        final_df = final_df.drop(columns=["VarNum_num"])
    else:
        final_df = final_df.sort_values(by=["Dataset", "Variable"])

    final_df = final_df.reset_index(drop=True)
    final_df.insert(0, "Order", range(1, len(final_df) + 1))

    return final_df


def build_datasets_spec_from_domains_config(mapping_df, config_df):
    if mapping_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Label", "Class", "Structure", "Key Variables", "Standard",
            "Repeat", "RefData"
        ])

    detected_datasets = sorted(mapping_df["SDTM Domain"].dropna().astype(str).str.upper().unique())
    expanded_cfg = expand_suppqual_to_supp_datasets(config_df, detected_datasets)

    ds_cols = [c for c in [
        "Dataset", "Dataset Label", "Class", "Structure", "Key Variables", "Repeat", "RefData"
    ] if c in expanded_cfg.columns]

    ds_df = expanded_cfg[ds_cols].drop_duplicates(subset=["Dataset"]).copy()
    ds_df = ds_df[ds_df["Dataset"].isin(detected_datasets)]

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
    if "Repeat" not in ds_df.columns:
        ds_df["Repeat"] = ""
    if "RefData" not in ds_df.columns:
        ds_df["RefData"] = ""

    ds_df["Standard"] = "SDTM"

    return ds_df[[
        "Dataset", "Label", "Class", "Structure", "Key Variables", "Standard",
        "Repeat", "RefData"
    ]].reset_index(drop=True)


# =========================================================
# Define sheet：Study info
# =========================================================
def extract_protocol_no_from_filename(file_name):
    """
    從檔名抓 Protocol No

    規則：
      1. sponsor_protocol no_eCRF schema XXX
      2. protocol no_eCRF schema XXX

    做法：
      - 先取 eCRF schema 前面的字串
      - 用 "_" 切開
      - 最後一段視為 protocol no
    """
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
            protocol_no,     # StudyName
            protocol_title,  # StudyDescription
            protocol_no,     # ProtocolName
            "SDTM-IG",
            std_ver,
            "en"
        ]
    })

    return define_df


def build_empty_codelists_sheet():
    return pd.DataFrame(columns=[
        "ID", "Name", "NCI Codelist Code", "Data Type", "Terminology",
        "Comment", "Order", "Term", "NCI Term Code", "Decoded Value"
    ])


def build_empty_dictionaries_sheet():
    return pd.DataFrame(columns=[
        "ID", "Name", "Data Type", "Dictionary", "Version"
    ])


def build_empty_trial_design_sheet(sheet_name):
    return pd.DataFrame()


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

        available_sheets = result["available_sheets"]
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
                # -------------------------------
                # 2.1 所有使用者輸入集中
                # -------------------------------
                st.markdown("### 2.1 Basic Information")

                version = st.selectbox(
                    "SDTM Version",
                    ["Version 3.3", "Version 3.4"],
                    key="sdtm_version_selector"
                )

                default_protocol_no = extract_protocol_no_from_filename(uploaded_file.name)

                col_a, col_b = st.columns(2)
                with col_a:
                    protocol_no = st.text_input(
                        "Protocol No",
                        value=default_protocol_no,
                        key="protocol_no"
                    )
                with col_b:
                    protocol_title = st.text_input(
                        "Protocol Title",
                        value="",
                        key="protocol_title"
                    )

                try:
                    raw_cfg_df, cfg_path = load_domains_config(version)
                    cfg_df = standardize_domains_config(raw_cfg_df)

                    st.success(f"✅ 已成功載入 config：{cfg_path}")

                    # 先 Datasets 再 Variables
                    st.markdown("### 2.2 Datasets SPEC")
                    datasets_spec_df = build_datasets_spec_from_domains_config(
                        mapping_df=mapping_df,
                        config_df=cfg_df
                    )
                    st.dataframe(datasets_spec_df, use_container_width=True)

                    st.markdown("### 2.3 Variables SPEC")
                    variables_spec_df = build_variables_spec_from_domains_config(
                        detail_df=detail_df,
                        config_df=cfg_df
                    )
                    st.dataframe(variables_spec_df, use_container_width=True)

                    st.markdown("### 2.4 Define / Codelists / Dictionaries / Trial Design")
                    define_df = build_define_sheet(
                        version=version,
                        protocol_no=protocol_no,
                        protocol_title=protocol_title
                    )
                    codelists_df = build_empty_codelists_sheet()
                    dictionaries_df = build_empty_dictionaries_sheet()

                    ta_df = build_empty_trial_design_sheet("TA")
                    te_df = build_empty_trial_design_sheet("TE")
                    ti_df = build_empty_trial_design_sheet("TI")
                    ts_df = build_empty_trial_design_sheet("TS")
                    tv_df = build_empty_trial_design_sheet("TV")

                    st.dataframe(define_df, use_container_width=True)
                    st.dataframe(codelists_df, use_container_width=True)
                    st.dataframe(dictionaries_df, use_container_width=True)

                    # 指定輸出順序
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
