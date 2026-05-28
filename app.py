import streamlit as st
import pandas as pd
import re
import os
import tempfile
from io import BytesIO

# pyreadstat 可能在部署環境尚未安裝，所以做保護
try:
    import pyreadstat
    HAS_PYREADSTAT = True
except Exception:
    HAS_PYREADSTAT = False

st.set_page_config(page_title="Auto SDTM SPEC", layout="wide")
st.title("Auto SDTM SPEC")


# =========================================================
# 基本工具
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
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_column(columns, required_keywords):
    """
    在欄位名稱中找最符合的欄位
    """
    for col in columns:
        upper_col = normalize_text(col)
        if all(k.upper() in upper_col for k in required_keywords):
            return col
    return None


def find_sheet_name(sheet_names, target_name):
    """
    忽略大小寫找 sheet name
    """
    for s in sheet_names:
        if str(s).strip().upper() == str(target_name).strip().upper():
            return s
    return None


# =========================================================
# Header 偵測
# =========================================================
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
    """
    manual_header_row_excel: Excel 看到的列號（1-based）
    回傳:
      df, detected_header_row_excel
    """
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


# =========================================================
# SoA：抓 CRF domain / sheet
# =========================================================
def extract_form_oids(series):
    """
    從 SoA 的 Form OID 抽出 domain
    """
    domains = set()

    for value in series.dropna():
        text = str(value).strip()
        if not text:
            continue

        # SoA domain 使用逗號 / 換行 / 分號 / slash 都切
        parts = re.split(r"[,\n;/]+", text)

        for part in parts:
            item = part.strip()
            if item:
                domains.add(item.upper())

    return domains


# =========================================================
# Source CRF Variable：優先抓 Field OID
# =========================================================
def find_source_variable_column(columns):
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

    # exact match
    for target in priority_exact:
        for col, norm_col in normalized_map.items():
            if norm_col == target:
                return col

    # fuzzy: FIELD + OID
    for col, norm_col in normalized_map.items():
        if "FIELD" in norm_col and "OID" in norm_col:
            return col

    # fallback: VARIABLE
    for col, norm_col in normalized_map.items():
        if "VARIABLE" in norm_col and "TARGET" not in norm_col and "SDTM" not in norm_col:
            return col

    return None


# =========================================================
# SDTM IG Target parsing
# =========================================================
def parse_sdtm_targets(value):
    """
    規則：
      - 只用分號 ; 和換行拆
      - 不用逗號和斜線拆
      - 支援：
          AE.AETERM
          VS.VSTESTCD="TEMP"
          DM.SEX='F'

    回傳：
      parsed_records: list of dict
      unparsed_tokens: list
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


# =========================================================
# Step 1：建立 mapping
# =========================================================
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
            df, detected_header_row = read_sheet_with_detected_header(
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

        for idx, row in df.iterrows():
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
                        "Unparsed Token": token,
                        "Excel Data Row": idx + 1 + detected_header_row
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


def process_uploaded_excel(file_bytes, manual_soa_header=None, common_domain_header=None):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    all_sheets = xls.sheet_names

    soa_sheet = find_sheet_name(all_sheets, "SoA")
    if soa_sheet is None:
        raise ValueError("找不到 SoA 分頁")

    soa_df, _ = read_sheet_with_detected_header(
        file_bytes=file_bytes,
        sheet_name=soa_sheet,
        keyword_groups=[["FORM", "OID"]],
        manual_header_row_excel=manual_soa_header
    )

    form_oid_col = find_column(soa_df.columns, ["FORM", "OID"])
    if form_oid_col is None:
        raise ValueError("SoA 分頁中找不到 Form OID 欄位")

    valid_domains = extract_form_oids(soa_df[form_oid_col])
    sheet_upper_map = {str(s).upper(): s for s in all_sheets}

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
# Step 2：SAS7BDAT config
# =========================================================
def read_uploaded_sas7bdat(uploaded_file):
    """
    pyreadstat 通常吃 path，比較穩的做法是先存 temp file
    """
    if uploaded_file is None:
        return pd.DataFrame()

    if not HAS_PYREADSTAT:
        return pd.DataFrame()

    suffix = ".sas7bdat"
    tmp_path = None

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        df, meta = pyreadstat.read_sas7bdat(tmp_path)
        df = normalize_columns(df)
        return df

    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def map_variable_config_columns(df):
    """
    將 variables config 欄位整理成 app 內部需要的標準欄位
    期待可能出現的欄位：Domain / Dataset, Name / Variable, Label, Type, Origin...
    """
    if df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment"
        ])

    cols = list(df.columns)

    dataset_col = find_column(cols, ["DATASET"]) or find_column(cols, ["DOMAIN"])
    variable_col = find_column(cols, ["VARIABLE"]) or find_column(cols, ["NAME"])
    label_col = find_column(cols, ["LABEL"])
    datatype_col = find_column(cols, ["DATA", "TYPE"]) or find_column(cols, ["TYPE"])
    codelist_col = find_column(cols, ["CODELIST"])
    origin_col = find_column(cols, ["ORIGIN"])
    source_col = find_column(cols, ["SOURCE"])
    pages_col = find_column(cols, ["PAGES"])
    method_col = find_column(cols, ["METHOD"])
    comment_col = find_column(cols, ["COMMENT"])

    out = pd.DataFrame()

    out["Dataset"] = df[dataset_col].astype(str).str.upper() if dataset_col else ""
    out["Variable"] = df[variable_col].astype(str).str.upper() if variable_col else ""
    out["Label"] = df[label_col] if label_col else ""
    out["Data Type"] = df[datatype_col] if datatype_col else ""
    out["Codelist"] = df[codelist_col] if codelist_col else ""
    out["Origin"] = df[origin_col] if origin_col else ""
    out["Source"] = df[source_col] if source_col else ""
    out["Pages"] = df[pages_col] if pages_col else ""
    out["Method"] = df[method_col] if method_col else ""
    out["Comment"] = df[comment_col] if comment_col else ""

    out = out[
        (out["Dataset"].astype(str).str.strip() != "") &
        (out["Variable"].astype(str).str.strip() != "")
    ].drop_duplicates()

    return out.reset_index(drop=True)


def map_dataset_config_columns(df):
    """
    將 datasets config 欄位整理成 app 內部需要的標準欄位
    期待可能出現的欄位：Domain / Dataset, Dlabel / Label, Structure, KeyVars...
    """
    if df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Label", "Class", "Structure", "Key Variables", "Standard"
        ])

    cols = list(df.columns)

    dataset_col = find_column(cols, ["DATASET"]) or find_column(cols, ["DOMAIN"])
    label_col = find_column(cols, ["DLABEL"]) or find_column(cols, ["LABEL"])
    class_col = find_column(cols, ["CLASS"])
    structure_col = find_column(cols, ["STRUCTURE"])
    keyvars_col = find_column(cols, ["KEY", "VAR"]) or find_column(cols, ["KEYVARIABLE"])
    standard_col = find_column(cols, ["STANDARD"])

    out = pd.DataFrame()

    out["Dataset"] = df[dataset_col].astype(str).str.upper() if dataset_col else ""
    out["Label"] = df[label_col] if label_col else ""
    out["Class"] = df[class_col] if class_col else ""
    out["Structure"] = df[structure_col] if structure_col else ""
    out["Key Variables"] = df[keyvars_col] if keyvars_col else ""
    out["Standard"] = df[standard_col] if standard_col else "SDTM"

    out = out[
        out["Dataset"].astype(str).str.strip() != ""
    ].drop_duplicates()

    return out.reset_index(drop=True)


def get_non_crf_from_config(detail_df, config_variables_df):
    """
    non-CRF = config variables - CRF mapping variables
    """
    if detail_df.empty or config_variables_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment"
        ])

    crf_pairs = set(
        zip(
            detail_df["SDTM Domain"].astype(str).str.upper(),
            detail_df["SDTM Variable"].astype(str).str.upper()
        )
    )

    detected_datasets = set(detail_df["SDTM Domain"].astype(str).str.upper())

    cfg = config_variables_df.copy()
    cfg["Dataset"] = cfg["Dataset"].astype(str).str.upper()
    cfg["Variable"] = cfg["Variable"].astype(str).str.upper()

    cfg = cfg[cfg["Dataset"].isin(detected_datasets)]

    cfg["pair"] = list(zip(cfg["Dataset"], cfg["Variable"]))
    non_crf = cfg[~cfg["pair"].isin(crf_pairs)].copy()

    if "pair" in non_crf.columns:
        non_crf = non_crf.drop(columns=["pair"])

    return non_crf.reset_index(drop=True)


def enrich_crf_variables_with_config(detail_df, config_variables_df):
    """
    將 Step 1 的 CRF mapping 用 config metadata enrich
    """
    if detail_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment"
        ])

    crf_df = detail_df.copy()
    crf_df["Dataset"] = crf_df["SDTM Domain"].astype(str).str.upper()
    crf_df["Variable"] = crf_df["SDTM Variable"].astype(str).str.upper()
    crf_df["Label"] = ""
    crf_df["Data Type"] = ""
    crf_df["Codelist"] = ""

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

    base_crf_df = crf_df[[
        "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment"
    ]].copy()

    if config_variables_df.empty:
        return base_crf_df.drop_duplicates().reset_index(drop=True)

    cfg = config_variables_df.copy()

    merged = base_crf_df.merge(
        cfg,
        on=["Dataset", "Variable"],
        how="left",
        suffixes=("", "_CFG")
    )

    for c in ["Label", "Data Type", "Codelist", "Origin", "Source", "Pages", "Method", "Comment"]:
        cfg_col = f"{c}_CFG"
        if cfg_col in merged.columns:
            merged[c] = merged.apply(
                lambda r: r[c] if str(r[c]).strip() != "" else (r[cfg_col] if pd.notna(r[cfg_col]) else ""),
                axis=1
            )
            merged = merged.drop(columns=[cfg_col])

    return merged.drop_duplicates().reset_index(drop=True)


def build_variables_spec(detail_df, config_variables_df):
    """
    Variables = CRF mapping + non-CRF from config
    """
    crf_part = enrich_crf_variables_with_config(detail_df, config_variables_df)
    non_crf_part = get_non_crf_from_config(detail_df, config_variables_df)

    combined = pd.concat([crf_part, non_crf_part], ignore_index=True)
    combined = combined.drop_duplicates().reset_index(drop=True)
    combined = combined.sort_values(by=["Dataset", "Variable"]).reset_index(drop=True)
    combined.insert(0, "Order", range(1, len(combined) + 1))

    return combined


def build_datasets_spec(variables_spec_df, config_datasets_df):
    if variables_spec_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Label", "Class", "Structure", "Key Variables", "Standard"
        ])

    datasets = sorted(variables_spec_df["Dataset"].dropna().astype(str).unique())

    rows = []
    for ds in datasets:
        ds_vars = variables_spec_df.loc[variables_spec_df["Dataset"] == ds, "Variable"].tolist()

        key_vars = []
        for candidate in ["STUDYID", "USUBJID", f"{ds}SEQ"]:
            if candidate in ds_vars:
                key_vars.append(candidate)

        rows.append({
            "Dataset": ds,
            "Label": "",
            "Class": "",
            "Structure": "",
            "Key Variables": " ".join(key_vars),
            "Standard": "SDTM"
        })

    ds_df = pd.DataFrame(rows)

    if not config_datasets_df.empty:
        cfg = config_datasets_df.copy()
        ds_df = ds_df.merge(cfg, on="Dataset", how="left", suffixes=("", "_CFG"))

        for c in ["Label", "Class", "Structure", "Key Variables", "Standard"]:
            cfg_col = f"{c}_CFG"
            if cfg_col in ds_df.columns:
                ds_df[c] = ds_df.apply(
                    lambda r: r[c] if str(r[c]).strip() != "" else (r[cfg_col] if pd.notna(r[cfg_col]) else ""),
                    axis=1
                )
                ds_df = ds_df.drop(columns=[cfg_col])

    return ds_df


def build_define_sheet():
    return pd.DataFrame([
        {"Attribute": "Standard", "Value": "SDTM"},
        {"Attribute": "StandardVersion", "Value": ""},
        {"Attribute": "StudyName", "Value": ""},
        {"Attribute": "ProtocolName", "Value": ""},
        {"Attribute": "Comment", "Value": ""}
    ])


def build_empty_codelists_sheet():
    return pd.DataFrame(columns=[
        "ID", "Name", "NCI Codelist Code", "Data Type", "Terminology",
        "Comment", "Order", "Term", "NCI Term Code", "Decoded Value"
    ])


def build_empty_dictionaries_sheet():
    return pd.DataFrame(columns=[
        "ID", "Name", "Data Type", "Dictionary", "Version"
    ])


def to_excel_bytes(sheet_dict):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheet_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0
