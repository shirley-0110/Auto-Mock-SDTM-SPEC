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
    代表只要某個儲存格同時包含該組關鍵字，就算找到
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
    自動或手動指定 header row後讀取 sheet
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

    return df, header_row_zero_based + 1  # 回傳 Excel 可讀列號（1-based）


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
      AE\nVIS
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
    從 SDTM IG Target 抓出 (domain, variable)

    支援範例：
      AE.AETERM
      DM.USUBJID
      AE.AEDECOD, AE.AETOXGR
      AE.AESTDTC; AE.AEENDTC
      DM.USUBJID / DM.SUBJID

    回傳：
      parsed_pairs: [(domain, variable), ...]
      unparsed_tokens: 不能解析的 token
    """
    parsed_pairs = []
    unparsed_tokens = []

    if pd.isna(value):
        return parsed_pairs, unparsed_tokens

    text = str(value).strip()
    if not text:
        return parsed_pairs, unparsed_tokens

    tokens = re.split(r"[,\n;/]+", text)

    for token in tokens:
        token = token.strip()
        if not token:
            continue

        # 抓像 DOMAIN.VARIABLE
        matches = re.findall(
            r"([A-Za-z][A-Za-z0-9]{0,7})\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)",
            token
        )

        if matches:
            for dom, var in matches:
                parsed_pairs.append((dom.upper(), var.upper()))
        else:
            unparsed_tokens.append(token)

    return parsed_pairs, unparsed_tokens


def build_sdtm_mapping(file_bytes, selected_crf_sheets, manual_sheet_headers=None):
    """
    從各 CRF sheet 的 SDTM IG Target 產生整份檔案的 SDTM domain/variable

    manual_sheet_headers:
      dict，例如 {"AE": 3, "DM": 2}
      表示某些 sheet 使用手動指定 header row（Excel 1-based）

    回傳：
      mapping_df
      source_detail_df
      sheet_without_target_col
      unparsed_records
      header_info_df
    """
    if manual_sheet_headers is None:
        manual_sheet_headers = {}

    records = []
    source_details = []
    sheet_without_target_col = []
    unparsed_records = []
    detected_header_info = []

    for sheet in selected_crf_sheets:
        try:
            manual_header = manual_sheet_headers.get(sheet)

            df, detected_header_row = read_sheet_with_detected_header(
                file_bytes=file_bytes,
                sheet_name=sheet,
                keyword_groups=[["SDTM", "TARGET"]],
                manual_header_row_excel=manual_header
            )

            detected_header_info.append({
                "Sheet": sheet,
                "Detected Header Row (Excel)": detected_header_row
            })

        except Exception as e:
            sheet_without_target_col.append(f"{sheet}（header 偵測失敗: {e}）")
            continue

        target_col = find_column(df.columns, ["SDTM", "TARGET"])

        if target_col is None:
            sheet_without_target_col.append(f"{sheet}（找不到 SDTM IG Target 欄位）")
            continue

        for idx, val in df[target_col].items():
            parsed_pairs, unparsed_tokens = parse_sdtm_targets(val)

            for dom, var in parsed_pairs:
                records.append({
                    "SDTM Domain": dom,
                    "SDTM Variable": var,
                    "Source CRF Sheet": sheet
                })

                source_details.append({
                    "Source CRF Sheet": sheet,
                    "Excel Data Row": idx + 1 + detected_header_row,
                    "SDTM IG Target Raw": val,
                    "SDTM Domain": dom,
                    "SDTM Variable": var
                })

            for token in unparsed_tokens:
                unparsed_records.append({
                    "Source CRF Sheet": sheet,
                    "Excel Data Row": idx + 1 + detected_header_row,
                    "SDTM IG Target Raw": val,
                    "Unparsed Token": token
                })

    if records:
        mapping_df = (
            pd.DataFrame(records)
            .drop_duplicates()
            .sort_values(by=["SDTM Domain", "SDTM Variable", "Source CRF Sheet"])
            .reset_index(drop=True)
        )
    else:
        mapping_df = pd.DataFrame(columns=["SDTM Domain", "SDTM Variable", "Source CRF Sheet"])

    if source_details:
        source_detail_df = pd.DataFrame(source_details).drop_duplicates().reset_index(drop=True)
    else:
        source_detail_df = pd.DataFrame(
            columns=["Source CRF Sheet", "Excel Data Row", "SDTM IG Target Raw", "SDTM Domain", "SDTM Variable"]
        )

    if detected_header_info:
        header_info_df = pd.DataFrame(detected_header_info).reset_index(drop=True)
    else:
        header_info_df = pd.DataFrame(columns=["Sheet", "Detected Header Row (Excel)"])

    return mapping_df, source_detail_df, sheet_without_target_col, unparsed_records, header_info_df


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
    summary_df["Variables"] = summary_df["SDTM Variable"].apply(lambda x: ", ".join(x))

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

        st.subheader("Excel 內所有 Sheets")
        st.write(all_sheets)

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

            st.subheader("SoA 偵測結果")
            c1, c2 = st.columns(2)

            with c1:
                st.write(f"SoA header row（Excel）: 第 {detected_soa_header} 列")

            with c2:
                st.write("SoA 欄位名稱：", list(soa_df.columns))

            form_oid_col = find_column(soa_df.columns, ["FORM", "OID"])

            if form_oid_col is None:
                st.error("SoA 分頁中找不到 Form OID 欄位")
            else:
                st.success(f"SoA 使用欄位：{form_oid_col}")

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

                c3, c4 = st.columns(2)

                with c3:
                    st.subheader("SoA 定義的 CRF Domains")
                    st.write(sorted(valid_domains))

                with c4:
                    st.subheader("可用的 CRF Sheets")
                    if available_sheets:
                        st.write(sorted(available_sheets))
                    else:
                        st.warning("SoA 沒有對應到任何實際存在的 sheet")

                if missing_sheets:
                    st.subheader("SoA 有，但 Excel 沒有的 Sheets")
                    st.warning(missing_sheets)

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
                mapping_df, source_detail_df, sheet_without_target_col, unparsed_records, header_info_df = build_sdtm_mapping(
                    file_bytes=file_bytes,
                    selected_crf_sheets=available_sheets,
                    manual_sheet_headers=manual_sheet_headers
                )

                st.subheader("各 Domain Sheet 偵測到的 Header Row")
                if not header_info_df.empty:
                    st.dataframe(header_info_df, use_container_width=True)
                else:
                    st.info("目前沒有可顯示的 header 偵測結果")

                # ---------------------------------------------
                # Step 4: SDTM summary
                # ---------------------------------------------
                st.subheader("整份檔案要呈現的 SDTM Domains / Variables")

                if mapping_df.empty:
                    st.warning("目前沒有從各 CRF sheet 的 SDTM IG Target 抓到可解析的 SDTM domain / variable")
                else:
                    summary_df = summarize_sdtm_mapping(mapping_df)
                    st.dataframe(summary_df, use_container_width=True)

                    with st.expander("查看 SDTM domain-variable 明細（含來源 sheet）", expanded=False):
                        st.dataframe(mapping_df, use_container_width=True)

                    with st.expander("查看來源明細（每列 SDTM IG Target 對應）", expanded=False):
                        st.dataframe(source_detail_df, use_container_width=True)

                # ---------------------------------------------
                # Step 5: 問題提示
                # ---------------------------------------------
                if sheet_without_target_col:
                    st.subheader("找不到 SDTM IG Target 欄位 / Header 偵測失敗的 Sheets")
                    st.warning(sheet_without_target_col)

                if unparsed_records:
                    st.subheader("無法解析的 SDTM IG Target 值")
                    st.dataframe(pd.DataFrame(unparsed_records), use_container_width=True)

                # ---------------------------------------------
                # Step 6: 單一 CRF sheet 預覽
                # ---------------------------------------------
                if available_sheets:
                    st.subheader("查看單一 CRF Sheet")

                    selected_sheet = st.selectbox(
                        "請選擇要查看的 CRF Domain Sheet",
                        options=sorted(available_sheets)
                    )

                    preview_manual_header = manual_sheet_headers.get(selected_sheet)

                    try:
                        preview_df, preview_header_row = read_sheet_with_detected_header(
                            file_bytes=file_bytes,
                            sheet_name=selected_sheet,
                            keyword_groups=[["SDTM", "TARGET"]],
                            manual_header_row_excel=preview_manual_header
                        )

                        st.write(f"{selected_sheet} header row（Excel）: 第 {preview_header_row} 列")
                        st.write(f"資料筆數：{len(preview_df)}")
                        st.write("欄位名稱：", list(preview_df.columns))
                        st.dataframe(preview_df, use_container_width=True)

                    except Exception as e:
                        st.error(f"讀取 {selected_sheet} 時發生錯誤：{e}")

    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
``
