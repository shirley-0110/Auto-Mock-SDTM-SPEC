import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="CRF → SDTM Target Viewer", layout="wide")
st.title("CRF → SDTM Target Viewer")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

# =========================================================
# 設定：Excel 變數名稱列
# 如果你的欄位名稱在第 2 列，header=1
# 如果之後有別的檔案不是第 2 列，再改這裡即可
# =========================================================
HEADER_ROW = 1


# =========================================================
# 工具函式
# =========================================================
def normalize_columns(df):
    """清理欄位名稱：去前後空白、換行、NBSP、多重空白"""
    cols = []
    for c in df.columns:
        c = str(c)
        c = c.replace("\n", " ")
        c = c.replace("\xa0", " ")
        c = re.sub(r"\s+", " ", c).strip()
        cols.append(c)
    df.columns = cols
    return df


def read_sheet(file_bytes, sheet_name, header_row=1):
    """讀單一 sheet，並清理欄位名稱"""
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
    df = normalize_columns(df)
    return df


def find_column(columns, required_keywords):
    """
    在欄位列表中模糊找欄位
    required_keywords: 例如 ["FORM", "OID"]
    """
    for col in columns:
        upper_col = col.upper()
        if all(k.upper() in upper_col for k in required_keywords):
            return col
    return None


def extract_form_oids(series):
    """
    從 SoA 的 Form OID 欄位抓出 CRF domain
    支援：
    - DM
    - AE, VIS
    - AE;VIS
    - AE\nVIS
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


def parse_sdtm_targets(value):
    """
    從 SDTM IG Target 欄位內容中抓出 (domain, variable)
    支援：
    - AE.AETERM
    - AE.AEDECOD, AE.AETOXGR
    - AE.AETERM; DM.USUBJID
    - 多個 target 混在同一格

    回傳:
    - parsed_pairs: [(domain, variable), ...]
    - unparsed_tokens: 無法辨識的 token
    """
    parsed_pairs = []
    unparsed_tokens = []

    if pd.isna(value):
        return parsed_pairs, unparsed_tokens

    text = str(value).strip()
    if not text:
        return parsed_pairs, unparsed_tokens

    # 先依常見分隔符拆開
    tokens = re.split(r"[,\n;/]+", text)

    for token in tokens:
        token = token.strip()
        if not token:
            continue

        # 找像 DOMAIN.VARIABLE 的格式
        # 例如 AE.AETERM / DM.USUBJID
        matches = re.findall(r"([A-Za-z][A-Za-z0-9]{0,7})\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)", token)

        if matches:
            for dom, var in matches:
                parsed_pairs.append((dom.upper(), var.upper()))
        else:
            unparsed_tokens.append(token)

    return parsed_pairs, unparsed_tokens


def build_sdtm_mapping(file_bytes, selected_crf_sheets, header_row=1):
    """
    根據 SoA 篩出的 CRF sheets，逐一讀取其 SDTM IG Target，
    建立整份檔案的 SDTM domain / variable 對應

    回傳：
    - mapping_df: 彙總表（SDTM Domain / SDTM Variable / Source CRF Sheet）
    - source_detail_df: 每個來源sheet與target原值的明細
    - sheet_without_target_col: 找不到 SDTM IG Target 欄位的 sheets
    - unparsed_records: 無法解析的 target 值
    """
    records = []
    source_details = []
    sheet_without_target_col = []
    unparsed_records = []

    for sheet in selected_crf_sheets:
        try:
            df = read_sheet(file_bytes, sheet, header_row=header_row)
        except Exception as e:
            sheet_without_target_col.append(f"{sheet}（讀取失敗: {e}）")
            continue

        target_col = find_column(df.columns, ["SDTM", "TARGET"])

        if target_col is None:
            sheet_without_target_col.append(sheet)
            continue

        # 一列一列抓 target
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
                    "Excel Row": idx + 1 + HEADER_ROW,  # 顯示接近實際Excel列號
                    "SDTM IG Target Raw": val,
                    "SDTM Domain": dom,
                    "SDTM Variable": var
                })

            for token in unparsed_tokens:
                unparsed_records.append({
                    "Source CRF Sheet": sheet,
                    "Excel Row": idx + 1 + HEADER_ROW,
                    "SDTM IG Target Raw": val,
                    "Unparsed Token": token
                })

    if records:
        mapping_df = pd.DataFrame(records).drop_duplicates().sort_values(
            by=["SDTM Domain", "SDTM Variable", "Source CRF Sheet"]
        )
    else:
        mapping_df = pd.DataFrame(columns=["SDTM Domain", "SDTM Variable", "Source CRF Sheet"])

    if source_details:
        source_detail_df = pd.DataFrame(source_details).drop_duplicates()
    else:
        source_detail_df = pd.DataFrame(
            columns=["Source CRF Sheet", "Excel Row", "SDTM IG Target Raw", "SDTM Domain", "SDTM Variable"]
        )

    return mapping_df, source_detail_df, sheet_without_target_col, unparsed_records


# =========================================================
# 主流程
# =========================================================
if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        xls = pd.ExcelFile(BytesIO(file_bytes))
        all_sheets = xls.sheet_names

        st.subheader("Excel 內所有 Sheets")
        st.write(all_sheets)

        # -----------------------------
        # Step 1: 讀 SoA
        # -----------------------------
        if "SoA" not in all_sheets:
            st.error("找不到 SoA 分頁")
        else:
            soa_df = read_sheet(file_bytes, "SoA", header_row=HEADER_ROW)

            st.subheader("SoA 欄位偵測")
            st.write(list(soa_df.columns))

            form_oid_col = find_column(soa_df.columns, ["FORM", "OID"])

            if form_oid_col is None:
                st.error("SoA 分頁中找不到 Form OID 欄位")
            else:
                st.success(f"SoA 使用欄位：{form_oid_col}")

                # -----------------------------
                # Step 2: 從 SoA 抓有效 CRF sheets
                # -----------------------------
                valid_domains = extract_form_oids(soa_df[form_oid_col])

                st.subheader("SoA 定義的 CRF Domains")
                st.write(sorted(valid_domains))

                sheet_upper_map = {s.upper(): s for s in all_sheets}

                available_sheets = [
                    sheet_upper_map[d] for d in valid_domains if d in sheet_upper_map
                ]

                missing_sheets = [
                    d for d in valid_domains if d not in sheet_upper_map
                ]

                c1, c2 = st.columns(2)

                with c1:
                    st.subheader("可顯示的 CRF Sheets")
                    if available_sheets:
                        st.write(sorted(available_sheets))
                    else:
                        st.warning("SoA 沒有對應到任何實際存在的 sheet")

                with c2:
                    st.subheader("SoA 有，但 Excel 沒有的 Sheets")
                    if missing_sheets:
                        st.warning(missing_sheets)
                    else:
                        st.success("沒有缺少的 sheet")

                # -----------------------------
                # Step 3: 建 SDTM domain/variable mapping
                # -----------------------------
                mapping_df, source_detail_df, sheet_without_target_col, unparsed_records = build_sdtm_mapping(
                    file_bytes=file_bytes,
                    selected_crf_sheets=available_sheets,
                    header_row=HEADER_ROW
                )

                st.subheader("整份檔案要呈現的 SDTM Domains / Variables")

                if mapping_df.empty:
                    st.warning("目前沒有從各 CRF sheet 的 SDTM IG Target 抓到可解析的 SDTM domain / variable")
                else:
                    # 方便看：每個 domain 彙總有哪些 variables
                    domain_summary = (
                        mapping_df.groupby("SDTM Domain")["SDTM Variable"]
                        .apply(lambda x: sorted(set(x)))
                        .reset_index()
                    )
                    domain_summary["SDTM Variable Count"] = domain_summary["SDTM Variable"].apply(len)
                    domain_summary["SDTM Variables"] = domain_summary["SDTM Variable"].apply(lambda x: ", ".join(x))

                    st.dataframe(
                        domain_summary[["SDTM Domain", "SDTM Variable Count", "SDTM Variables"]],
                        use_container_width=True
                    )

                    # 明細
                    with st.expander("查看 SDTM domain-variable 明細（含來源 sheet）", expanded=False):
                        st.dataframe(mapping_df, use_container_width=True)

                    with st.expander("查看來源明細（每列 SDTM IG Target 對應）", expanded=False):
                        st.dataframe(source_detail_df, use_container_width=True)

                # -----------------------------
                # Step 4: 顯示有問題的 sheets / targets
                # -----------------------------
                if sheet_without_target_col:
                    st.subheader("缺少 SDTM IG Target 欄位的 CRF Sheets")
                    st.warning(sheet_without_target_col)

                if unparsed_records:
                    st.subheader("無法解析的 SDTM IG Target 值")
                    st.dataframe(pd.DataFrame(unparsed_records), use_container_width=True)

                # -----------------------------
                # Step 5: 使用者只看 SoA 篩選後的 sheet
                # -----------------------------
                if available_sheets:
                    st.subheader("查看單一 CRF Sheet")
                    selected_sheet = st.selectbox("請選擇 CRF Domain", sorted(available_sheets))

                    df_selected = read_sheet(file_bytes, selected_sheet, header_row=HEADER_ROW)

                    st.write(f"資料筆數：{len(df_selected)}")
                    st.write("欄位名稱：", list(df_selected.columns))
                    st.dataframe(df_selected, use_container_width=True)

    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
