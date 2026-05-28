

import streamlit as st
import pandas as pd
import re

st.title("Excel CRF Sheet Viewer")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

def extract_form_oids(series):
    """
    將 SoA 的 Form OID 欄位內容整理成 domain 清單
    支援：
    - 一格一個值：DM / AE / VIS
    - 一格多個值：DM, AE, VIS
    - 用逗號、分號、換行分隔
    """
    domains = set()

    for value in series.dropna():
        text = str(value).strip()
        if not text:
            continue

        # 依逗號、分號、換行切開
        parts = re.split(r"[,\n;]+", text)

        for part in parts:
            item = part.strip()
            if item:
                domains.add(item.upper())

    return domains


if uploaded_file is not None:
    try:
        # 先取得所有 sheet 名稱
        xls = pd.ExcelFile(uploaded_file)
        all_sheets = xls.sheet_names

        # 檢查是否有 SoA
        if "SoA" not in all_sheets:
            st.error("找不到 SoA 分頁")
        else:
            # 讀 SoA
            soa_df = pd.read_excel(uploaded_file, sheet_name="SoA")

            # 檢查是否有 Form OID 欄位
            if "Form OID" not in soa_df.columns:
                st.error("SoA 分頁中找不到 'Form OID' 欄位")
            else:
                # 從 Form OID 擷取 domain 名稱
                valid_domains = extract_form_oids(soa_df["Form OID"])

                # 只保留 Excel 裡真正存在的 sheet
                available_sheets = [
                    sheet for sheet in all_sheets
                    if sheet.upper() in valid_domains
                ]

                if not available_sheets:
                    st.warning("SoA 的 Form OID 沒有對應到任何實際存在的 sheet")
                else:
                    st.success("已依 SoA 的 Form OID 篩選可顯示的 sheet")

                    # 只顯示這些 sheet 給使用者選
                    selected_sheet = st.selectbox("請選擇要查看的 CRF sheet", available_sheets)

                    # 讀取選到的 sheet
                    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

                    st.subheader(f"Sheet：{selected_sheet}")
                    st.write(f"資料筆數：{len(df)}")
                    st.write("欄位名稱：", df.columns.tolist())
                    st.dataframe(df)

    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")


