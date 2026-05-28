import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="CRF → SDTM Mapping / SPEC Generator", layout="wide")
st.title("CRF → SDTM Mapping / SPEC Generator")


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


def build_sdtm_mapping(file_bytes, selected_crf_sheets, manual_sheet_headers=None):
    if manual_sheet_headers is None:
        manual_sheet_headers = {}

    mapping_records = []
    detail_records = []
    sheet_errors = []
    unparsed_records = []

    for sheet in selected_crf_sheets:
        try:
            manual_header = manual_sheet_headers.get(sheet)

            df, detected_header_row = read_sheet_with_detected_header(
                file_bytes=file_bytes,
                sheet_name=sheet,
                keyword_groups=[["SDTM", "TARGET"]],
                manual_header_row_excel=manual_header
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
            .sort_values(by=[
                "SDTM Domain",
                "SDTM Variable",
                "Source CRF Sheet",
                "Source CRF Variable",
                "Assign Value"
            ])
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
# SPEC mode
# =========================================================
def get_default_non_crf_rows(datasets):
    """
    預設建議值，可在 UI 直接修改
    """
    rows = []

    for ds in sorted(set(datasets)):
        if ds == "DM":
            defaults = [
                ("STUDYID", "Study Identifier", "text", "", "Assigned", "", "", "", ""),
                ("DOMAIN", "Domain Abbreviation", "text", "", "Assigned", "", "", "", ""),
                ("USUBJID", "Unique Subject Identifier", "text", "", "Derived", "", "", "", ""),
            ]
        else:
            defaults = [
                ("STUDYID", "Study Identifier", "text", "", "Assigned", "", "", "", ""),
                ("DOMAIN", "Domain Abbreviation", "text", "", "Assigned", "", "", "", ""),
                ("USUBJID", "Unique Subject Identifier", "text", "", "Derived", "", "", "", ""),
                (f"{ds}SEQ", f"Sequence Number", "integer", "", "Derived", "", "", "", ""),
            ]

        for var, label, dtype, codelist, origin, source, pages, method, comment in defaults:
            rows.append({
                "Dataset": ds,
                "Variable": var,
                "Label": label,
                "Data Type": dtype,
                "Codelist": codelist,
                "Origin": origin,
                "Source": source,
                "Pages": pages,
                "Method": method,
                "Comment": comment
            })

    return pd.DataFrame(rows)


def build_variables_spec(detail_df, non_crf_df=None):
    """
    將 mapping detail 轉成 Variables SPEC 初版
    """
    if detail_df.empty:
        crf_df = pd.DataFrame(columns=[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment"
        ])
    else:
        crf_df = detail_df.copy()
        crf_df["Dataset"] = crf_df["SDTM Domain"]
        crf_df["Variable"] = crf_df["SDTM Variable"]
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

        crf_df = crf_df[[
            "Dataset", "Variable", "Label", "Data Type", "Codelist",
            "Origin", "Source", "Pages", "Method", "Comment"
        ]]

    if non_crf_df is not None and not non_crf_df.empty:
        combined = pd.concat([crf_df, non_crf_df], ignore_index=True)
    else:
        combined = crf_df.copy()

    combined = combined.drop_duplicates().reset_index(drop=True)
    combined = combined.sort_values(by=["Dataset", "Variable"]).reset_index(drop=True)
    combined.insert(0, "Order", range(1, len(combined) + 1))

    return combined


def build_datasets_spec(variables_spec_df):
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

    return pd.DataFrame(rows)


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

    output.seek(0)
    return output.getvalue()


# =========================================================
# 共用：從 Excel 建立 CRF mapping
# =========================================================
def process_uploaded_excel(file_bytes, all_sheets, manual_soa_header=None, manual_sheet_headers=None):
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
        manual_sheet_headers=manual_sheet_headers or {}
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
# UI
# =========================================================
mode = st.sidebar.selectbox(
    "功能選擇",
    ["CRF → SDTM Mapping", "SDTM SPEC Generator"]
)

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        xls = pd.ExcelFile(BytesIO(file_bytes))
        all_sheets = xls.sheet_names

        # Sidebar：header override
        st.sidebar.header("Header Override（選填）")

        use_manual_soa_header = st.sidebar.checkbox("手動指定 SoA header row")
        manual_soa_header = None

        if use_manual_soa_header:
            manual_soa_header = st.sidebar.number_input(
                "SoA header 在 Excel 第幾列？",
                min_value=1,
                value=2,
                step=1
            )

        manual_sheet_headers = {}

        # 如果要先猜一次 available sheets 來讓 sidebar 選 override，這裡先做個輕量處理
        try:
            tmp_soa_df, _ = read_sheet_with_detected_header(
                file_bytes=file_bytes,
                sheet_name="SoA",
                keyword_groups=[["FORM", "OID"]],
                manual_header_row_excel=manual_soa_header
            )
            tmp_form_oid_col = find_column(tmp_soa_df.columns, ["FORM", "OID"])
            if tmp_form_oid_col is not None:
                tmp_valid_domains = extract_form_oids(tmp_soa_df[tmp_form_oid_col])
                tmp_sheet_upper_map = {s.upper(): s for s in all_sheets}
                tmp_available_sheets = [
                    tmp_sheet_upper_map[d] for d in tmp_valid_domains if d in tmp_sheet_upper_map
                ]
            else:
                tmp_available_sheets = []
        except Exception:
            tmp_available_sheets = []

        if tmp_available_sheets:
            st.sidebar.subheader("手動指定個別 Domain Sheet Header")
            selected_override_sheets = st.sidebar.multiselect(
                "選擇需要手動指定 header 的 sheet",
                options=sorted(tmp_available_sheets)
            )

            for sh in selected_override_sheets:
                manual_sheet_headers[sh] = st.sidebar.number_input(
                    f"{sh} header 在 Excel 第幾列？",
                    min_value=1,
                    value=2,
                    step=1,
                    key=f"header_override_{sh}"
                )

        # 執行共用 mapping 流程
        result = process_uploaded_excel(
            file_bytes=file_bytes,
            all_sheets=all_sheets,
            manual_soa_header=manual_soa_header,
            manual_sheet_headers=manual_sheet_headers
        )

        available_sheets = result["available_sheets"]
        missing_sheets = result["missing_sheets"]
        mapping_df = result["mapping_df"]
        detail_df = result["detail_df"]
        sheet_errors = result["sheet_errors"]
        unparsed_records = result["unparsed_records"]

        if missing_sheets:
            st.warning(f"SoA 有但 Excel 沒有的 Sheets：{missing_sheets}")

        if mode == "CRF → SDTM Mapping":
            st.subheader("整份檔案要呈現的 SDTM Domains / Variables")

            if mapping_df.empty:
                st.warning("目前沒有從各 CRF sheet 的 SDTM IG Target 抓到可解析的 SDTM domain / variable")
            else:
                summary_df = summarize_sdtm_mapping(mapping_df)
                st.dataframe(summary_df, use_container_width=True)

            st.subheader("SDTM Mapping 明細")
            if detail_df.empty:
                st.info("目前沒有可顯示的明細")
            else:
                st.dataframe(detail_df, use_container_width=True)

            if sheet_errors:
                clean_sheets = sorted(set(sheet_errors))
                st.subheader("無法處理的 Sheets")
                st.warning(f"header 偵測失敗，無法自動判斷 header row: {clean_sheets}")

            if unparsed_records:
                st.subheader("無法解析的 SDTM IG Target 值")
                st.dataframe(pd.DataFrame(unparsed_records), use_container_width=True)

        elif mode == "SDTM SPEC Generator":
            st.subheader("Step 1｜CRF → SDTM Mapping 結果")

            if mapping_df.empty:
                st.warning("目前沒有可用的 CRF → SDTM mapping，無法進一步建立 SPEC")
            else:
                summary_df = summarize_sdtm_mapping(mapping_df)
                st.dataframe(summary_df, use_container_width=True)

                with st.expander("查看 SDTM Mapping 明細"):
                    st.dataframe(detail_df, use_container_width=True)

            if sheet_errors:
                clean_sheets = sorted(set(sheet_errors))
                st.warning(f"header 偵測失敗，無法自動判斷 header row: {clean_sheets}")

            if unparsed_records:
                with st.expander("查看無法解析的 SDTM IG Target 值"):
                    st.dataframe(pd.DataFrame(unparsed_records), use_container_width=True)

            st.subheader("Step 2｜補充 non-CRF Variables（可直接編輯）")

            detected_datasets = sorted(mapping_df["SDTM Domain"].dropna().unique()) if not mapping_df.empty else []
            seed_non_crf_df = get_default_non_crf_rows(detected_datasets)

            edited_non_crf_df = st.data_editor(
                seed_non_crf_df,
                num_rows="dynamic",
                use_container_width=True,
                key="non_crf_editor"
            )

            st.subheader("Step 3｜Variables SPEC（初版）")

            variables_spec_df = build_variables_spec(detail_df, edited_non_crf_df)
            st.dataframe(variables_spec_df, use_container_width=True)

            st.subheader("Step 4｜Datasets SPEC（初版）")
            datasets_spec_df = build_datasets_spec(variables_spec_df)
            datasets_spec_df = st.data_editor(
                datasets_spec_df,
                num_rows="dynamic",
                use_container_width=True,
                key="datasets_spec_editor"
            )

            st.subheader("Step 5｜Define / Codelists / Dictionaries（模板）")
            define_df = st.data_editor(
                build_define_sheet(),
                num_rows="dynamic",
                use_container_width=True,
                key="define_editor"
            )

            codelists_df = st.data_editor(
                build_empty_codelists_sheet(),
                num_rows="dynamic",
                use_container_width=True,
                key="codelists_editor"
            )

            dictionaries_df = st.data_editor(
                build_empty_dictionaries_sheet(),
                num_rows="dynamic",
                use_container_width=True,
                key="dictionaries_editor"
            )

            export_sheets = {
                "Define": define_df,
                "Datasets": datasets_spec_df,
                "Variables": variables_spec_df,
                "Codelists": codelists_df,
                "Dictionaries": dictionaries_df
            }

            excel_bytes = to_excel_bytes(export_sheets)

            st.download_button(
                label="下載 SDTM SPEC Excel",
                data=excel_bytes,
                file_name="SDTM_SPEC_Draft.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
