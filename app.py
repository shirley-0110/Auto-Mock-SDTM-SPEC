import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="CRF → SDTM Target Viewer", layout="wide")
st.title("CRF → SDTM Target Viewer")


# =========================================================
# 基本工具函式
# =========================================================
def normalize_text(x):
    """清理字串，方便欄位 / 儲存格比對"""
    if pd.isna(x):
        return ""
    x = str(x)
    x = x.replace("\n", " ")
    x = x.replace("\r", " ")
    x = x.replace("\xa0", " ")   # non-breaking space
    x = re.sub(r"\s+", " ", x)
    return x.strip().upper()


def normalize_columns(df):
    """清理 DataFrame 欄位名稱"""
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
    """
    檢查某一列是否包含目標欄位名稱
    keyword_groups 範例:
      [["FORM", "OID"]]
      [["SDTM", "TARGET"]]
    """
    cells = [normalize_text(v) for v in row_values]

    for cell in cells:
        for group in keyword_groups:
            if all(k.upper() in cell for k in group):
                return True
    return False


def detect_header_row(file_bytes, sheet_name, keyword_groups, max_scan_rows=30):
    """
    掃描前幾列，找出 header row（0-based）
    若找不到則回傳 None
    """
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
    自動或手動指定 header row 後讀取 sheet
    manual_header_row_excel: Excel 中第幾列是 header（1-based）
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

    return df, header_row_zero_based + 1  # 回傳 Excel 列號（1-based）


def find_column(columns, required_keywords):
    """
    在欄位名稱中找最符合的欄位
    例如:
      find_column(df.columns, ["FORM", "OID"])
      find_column(df.columns, ["SDTM", "TARGET"])
    """
    for col in columns:
        upper_col = normalize_text(col)
        if all(k.upper() in upper_col for k in required_keywords):
            return col
    return None


def find_source_variable_column(columns):
    """
    找來源 CRF variable 欄位
    優先找：
      - Field OID
      - FieldOID
      - Field OID Name
    再 fallback 到 Variable 類欄位
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

    # 先 exact match
    for target in priority_exact:
        for col, norm_col in normalized_map.items():
            if norm_col == target:
                return col

    # 再 fuzzy: 含 FIELD 與 OID
    for col, norm_col in normalized_map.items():
        if "FIELD" in norm_col and "OID" in norm_col:
            return col

    # 再 fuzzy: 含 VARIABLE 但排除 TARGET / SDTM
    for col, norm_col in normalized_map.items():
        if "VARIABLE" in norm_col and "TARGET" not in norm_col and "SDTM" not in norm_col:
            return col

    return None


# =========================================================
# SoA：抓 CRF domain / sheet
# =========================================================
def extract_form_oids(series):
    """
    從 SoA 的 Form OID 欄位抽出 domain
    支援：
      DM
      AE, VIS
      AE;VIS
      AE/VIS
      AE\\nVIS
    """
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
    解析 SDTM IG Target

    規則：
      - 只用分號 ; 和換行當作多 target 分隔
      - 不用逗號 , 和斜線 / 分隔
      - 支援：
          AE.AETERM
          VS.VSTESTCD="TEMP"
          DM.SEX='F'

    回傳：
      parsed_records: list of dict
        {
          "SDTM Domain": ...,
          "SDTM Variable": ...,
          "Assign Value": ...
        }
      unparsed_tokens: list
    """
    parsed_records = []
    unparsed_tokens = []

    if pd.isna(value):
        return parsed_records, unparsed_tokens

    text = str(value).strip()
    if not text:
        return parsed_records, unparsed_tokens

    # 只用分號與換行切
    tokens = re.split(r"[;\n]+", text)

    # 支援：
    #   VS.VSTESTCD
    #   VS.VSTESTCD="TEMP"
    #   VS.VSTESTCD='TEMP'
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
    """
    從各 CRF sheet 的 SDTM IG Target 產生整份檔案的 SDTM domain/variable

    回傳：
      mapping_df
      detail_df
      sheet_errors
      unparsed_records
    """
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

        except Exception as e:
            sheet_errors.append(f"{sheet}（header 偵測失敗: {e}）")
            continue

        target_col = find_column(df.columns, ["SDTM", "TARGET"])
        if target_col is None:
            sheet_errors.append(f"{sheet}（找不到 SDTM IG Target 欄位）")
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
    """
    將 mapping_df 整成 domain summary
    """
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
# UI
# =========================================================
uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        xls = pd.ExcelFile(BytesIO(file_bytes))
        all_sheets = xls.sheet_names

        # ---------------------------------------------
        # Sidebar：SoA header override
        # ---------------------------------------------
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

        # ---------------------------------------------
        # Step 1: 讀取 SoA（自動抓 Form OID 所在列）
        # ---------------------------------------------
        if "SoA" not in all_sheets:
            st.error("找不到 SoA 分頁")
        else:
            soa_df, detected_soa_header = read_sheet_with_detected_header(
                file_bytes=file_bytes,
                sheet_name="SoA",
                keyword_groups=[["FORM", "OID"]],
                manual_header_row_excel=manual_soa_header
            )

            form_oid_col = find_column(soa_df.columns, ["FORM", "OID"])

            if form_oid_col is None:
                st.error("SoA 分頁中找不到 Form OID 欄位")
            else:
                # ---------------------------------------------
                # Step 2: 從 SoA 抓 CRF domains
                # ---------------------------------------------
                valid_domains = extract_form_oids(soa_df[form_oid_col])

                sheet_upper_map = {s.upper(): s for s in all_sheets}

                available_sheets = [
                    sheet_upper_map[d] for d in valid_domains if d in sheet_upper_map
                ]

                missing_sheets = [
                    d for d in valid_domains if d not in sheet_upper_map
                ]

                if missing_sheets:
                    st.warning(f"SoA 有但 Excel 沒有的 Sheets：{missing_sheets}")

                # ---------------------------------------------
                # Sidebar：手動指定個別 sheet 的 header
                # ---------------------------------------------
                manual_sheet_headers = {}

                if available_sheets:
                    st.sidebar.subheader("手動指定個別 Domain Sheet Header")
                    selected_override_sheets = st.sidebar.multiselect(
                        "選擇需要手動指定 header 的 sheet",
                        options=sorted(available_sheets)
                    )

                    for sh in selected_override_sheets:
                        manual_sheet_headers[sh] = st.sidebar.number_input(
                            f"{sh} header 在 Excel 第幾列？",
                            min_value=1,
                            value=2,
                            step=1,
                            key=f"header_override_{sh}"
                        )

                # ---------------------------------------------
                # Step 3: 從各 sheet 的 SDTM IG Target 抓 SDTM domain/variable
                # ---------------------------------------------
                mapping_df, detail_df, sheet_errors, unparsed_records = build_sdtm_mapping(
                    file_bytes=file_bytes,
                    selected_crf_sheets=available_sheets,
                    manual_sheet_headers=manual_sheet_headers
                )

                # ---------------------------------------------
                # Step 4: SDTM summary
                # ---------------------------------------------
                st.subheader("整份檔案要呈現的 SDTM Domains / Variables")

                if mapping_df.empty:
                    st.warning("目前沒有從各 CRF sheet 的 SDTM IG Target 抓到可解析的 SDTM domain / variable")
                else:
                    summary_df = summarize_sdtm_mapping(mapping_df)
                    st.dataframe(summary_df, use_container_width=True)

                # ---------------------------------------------
                # Step 5: 整合後明細表
                # ---------------------------------------------
                st.subheader("SDTM Mapping 明細")

                if detail_df.empty:
                    st.info("目前沒有可顯示的明細")
                else:
                    st.dataframe(detail_df, use_container_width=True)

                # ---------------------------------------------
                # Step 6: 問題提示
                # ---------------------------------------------
                if sheet_errors:
                    # 把 sheet 名稱整理乾淨（去掉括號後面的說明）
                    clean_sheets = []
                
                    for err in sheet_errors:
                        # 只抓 sheet name（例如 "VS3（header 偵測失敗...）" → "VS3"）
                        sheet_name = re.split(r"（|\(", err)[0].strip()
                        clean_sheets.append(sheet_name)

                    # 去重 + 排序
                    clean_sheets = sorted(set(clean_sheets))

                    st.subheader("無法處理的 Sheets")

                    # ✅ 黃色警示盒 + 整體訊息
                    st.warning(
                        f"header 偵測失敗，無法自動判斷 header row: {clean_sheets}"
                    )


                if unparsed_records:
                    st.subheader("無法解析的 SDTM IG Target 值")
                    st.dataframe(pd.DataFrame(unparsed_records), use_container_width=True)

    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
