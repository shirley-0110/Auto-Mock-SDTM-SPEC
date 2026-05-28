import streamlit as st
import pandas as pd
import re
from io import BytesIO

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
# Reference SPEC 讀取 / 比對
# =========================================================
def load_reference_sheet(reference_file, sheet_name):
    try:
        xls = pd.ExcelFile(reference_file)
        if sheet_name not in xls.sheet_names:
            return pd.DataFrame()
        df = pd.read_excel(reference_file, sheet_name=sheet_name)
        df = normalize_columns(df)
        return df
    except Exception:
        return pd.DataFrame()


def get_non_crf_from_reference(detail_df, ref_variables_df):
    """
    從 reference Variables sheet 補 non-CRF variables：
    - 只保留目前 mapping 有出現的 datasets
    - 排除已經在 CRF mapping detail 裡出現的 (Dataset, Variable)
    """
    if detail_df.empty or ref_variables_df.empty:
        return pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment"
        ])

    if "Dataset" not in ref_variables_df.columns or "Variable" not in ref_variables_df.columns:
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

    ref = ref_variables_df.copy()
    ref["Dataset"] = ref["Dataset"].astype(str).str.upper()
    ref["Variable"] = ref["Variable"].astype(str).str.upper()

    ref = ref[ref["Dataset"].isin(detected_datasets)]

    ref["pair"] = list(zip(ref["Dataset"], ref["Variable"]))
    non_crf = ref[~ref["pair"].isin(crf_pairs)].copy()

    if "pair" in non_crf.columns:
        non_crf = non_crf.drop(columns=["pair"])

    for col in ["Label", "Data Type", "Codelist", "Origin", "Source", "Pages", "Method", "Comment"]:
        if col not in non_crf.columns:
            non_crf[col] = ""

    return non_crf[[
        "Dataset", "Variable", "Label", "Data Type", "Codelist",
        "Origin", "Source", "Pages", "Method", "Comment"
    ]].drop_duplicates().reset_index(drop=True)


def enrich_crf_variables_with_reference(detail_df, ref_variables_df):
    """
    若 reference Variables sheet 有同 Dataset / Variable，帶回 metadata
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

    if ref_variables_df.empty:
        return base_crf_df.drop_duplicates().reset_index(drop=True)

    if "Dataset" not in ref_variables_df.columns or "Variable" not in ref_variables_df.columns:
        return base_crf_df.drop_duplicates().reset_index(drop=True)

    ref = ref_variables_df.copy()
    ref["Dataset"] = ref["Dataset"].astype(str).str.upper()
    ref["Variable"] = ref["Variable"].astype(str).str.upper()

    merge_cols = ["Dataset", "Variable"]
    ref_meta_cols = []
    for c in ["Label", "Data Type", "Codelist", "Origin", "Source", "Pages", "Method", "Comment"]:
        if c in ref.columns:
            ref_meta_cols.append(c)

    if not ref_meta_cols:
        return base_crf_df.drop_duplicates().reset_index(drop=True)

    ref_for_merge = ref[merge_cols + ref_meta_cols].drop_duplicates()

    merged = base_crf_df.merge(
        ref_for_merge,
        on=["Dataset", "Variable"],
        how="left",
        suffixes=("", "_REF")
    )

    for c in ref_meta_cols:
        ref_col = f"{c}_REF"
        if ref_col in merged.columns:
            merged[c] = merged.apply(
                lambda r: r[c] if str(r[c]).strip() != "" else (r[ref_col] if pd.notna(r[ref_col]) else ""),
                axis=1
            )
            merged = merged.drop(columns=[ref_col])

    return merged.drop_duplicates().reset_index(drop=True)


def build_datasets_spec(variables_spec_df, ref_datasets_df=None):
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

    if ref_datasets_df is not None and not ref_datasets_df.empty and "Dataset" in ref_datasets_df.columns:
        ref = ref_datasets_df.copy()
        ref["Dataset"] = ref["Dataset"].astype(str).str.upper()

        keep_cols = ["Dataset"]
        for c in ["Label", "Class", "Structure", "Key Variables", "Standard"]:
            if c in ref.columns:
                keep_cols.append(c)

        ref = ref[keep_cols].drop_duplicates()

        merged = ds_df.merge(ref, on="Dataset", how="left", suffixes=("", "_REF"))

        for c in ["Label", "Class", "Structure", "Key Variables", "Standard"]:
            ref_col = f"{c}_REF"
            if ref_col in merged.columns:
                merged[c] = merged.apply(
                    lambda r: r[c] if str(r[c]).strip() != "" else (r[ref_col] if pd.notna(r[ref_col]) else ""),
                    axis=1
                )
                merged = merged.drop(columns=[ref_col])

        ds_df = merged

    return ds_df


def build_define_sheet(ref_define_df=None):
    if ref_define_df is not None and not ref_define_df.empty:
        return ref_define_df.copy()

    return pd.DataFrame([
        {"Attribute": "Standard", "Value": "SDTM"},
        {"Attribute": "StandardVersion", "Value": ""},
        {"Attribute": "StudyName", "Value": ""},
        {"Attribute": "ProtocolName", "Value": ""},
        {"Attribute": "Comment", "Value": ""}
    ])


def build_empty_codelists_sheet(ref_codelists_df=None):
    if ref_codelists_df is not None and not ref_codelists_df.empty:
        return ref_codelists_df.copy()

    return pd.DataFrame(columns=[
        "ID", "Name", "NCI Codelist Code", "Data Type", "Terminology",
        "Comment", "Order", "Term", "NCI Term Code", "Decoded Value"
    ])


def build_empty_dictionaries_sheet(ref_dictionaries_df=None):
    if ref_dictionaries_df is not None and not ref_dictionaries_df.empty:
        return ref_dictionaries_df.copy()

    return pd.DataFrame(columns=[
        "ID", "Name", "Data Type", "Dictionary", "Version"
    ])


def build_variables_spec(detail_df, non_crf_df=None, ref_variables_df=None):
    crf_df = enrich_crf_variables_with_reference(detail_df, ref_variables_df)

    if non_crf_df is not None and not non_crf_df.empty:
        combined = pd.concat([crf_df, non_crf_df], ignore_index=True)
    else:
        combined = crf_df.copy()

    combined = combined.drop_duplicates().reset_index(drop=True)
    combined = combined.sort_values(by=["Dataset", "Variable"]).reset_index(drop=True)
    combined.insert(0, "Order", range(1, len(combined) + 1))

    return combined


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
# 主流程 UI
# =========================================================
uploaded_file = st.file_uploader("請上傳 CRF Mapping Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    current_upload_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("current_upload_key") != current_upload_key:
        st.session_state["current_upload_key"] = current_upload_key
        st.session_state["run_step2"] = False

    try:
        file_bytes = uploaded_file.read()
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

        result = process_uploaded_excel(
            file_bytes=file_bytes,
            all_sheets=all_sheets,
            manual_soa_header=manual_soa_header,
            common_domain_header=common_domain_header
        )

        available_sheets = result["available_sheets"]
        missing_sheets = result["missing_sheets"]
        mapping_df = result["mapping_df"]
        detail_df = result["detail_df"]
        sheet_errors = result["sheet_errors"]
        unparsed_records = result["unparsed_records"]

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
        # Step 2 開關：使用者決定是否執行
        # -------------------------------------------------
        def trigger_step2():
            st.session_state["run_step2"] = True

        st.button(
            "▶ 執行 Step 2：SPEC Generator",
            type="primary",
            on_click=trigger_step2
        )

        # -------------------------------------------------
        # Step 2：SPEC Generator
        # -------------------------------------------------
        if st.session_state.get("run_step2", False):
            st.markdown("## Step 2｜SPEC Generator")

            if mapping_df.empty:
                st.warning("目前沒有可用的 CRF → SDTM mapping，無法建立 SPEC")
            else:
                # -------------------------------
                # 2.1 選擇 SDTM Version
                # -------------------------------
                st.markdown("### 2.1 選擇 SDTM Version")

                version = st.selectbox(
                    "請選擇 SDTM Version",
                    ["Version 3.3", "Version 3.4"]
                )

                # -------------------------------
                # 2.2 組 config 路徑
                # -------------------------------
                BASE_PATH = r"Y:\BS Files\CDISC\04. SDTM"

                config_base_path = f"{BASE_PATH}\\{version}"

                st.write("使用的 config 路徑：")
                st.code(config_base_path)


                # -------------------------------
                # 2.3 嘗試讀取 SAS config
                # -------------------------------
                st.markdown("### 2.2 載入 SAS Config")

                config_loaded = False
                var_cfg_df = pd.DataFrame()
                ds_cfg_df = pd.DataFrame()

                try:
                    import pyreadstat

                    # 👉 這裡你可以改成你實際檔名
                    ds_file_path  = f"{config_base_path}\\Datasets\domains.sas7bdat"
                    ds_cfg_df, _ = pyreadstat.read_sas7bdat(ds_file_path)

                    config_loaded = True

                except Exception as e:
                    st.warning(f"⚠ 讀取 config 失敗：{e}")


    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
