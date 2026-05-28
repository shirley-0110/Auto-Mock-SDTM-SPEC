
import streamlit as st
import pandas as pd
import re

st.title("CRF Excel Viewer")

uploaded_file = st.file_uploader("請上傳 CRF Schema 檔案", type=["xlsx", "xls"])


# ✅ function：從 Form OID 抽 domain
def extract_form_oids(series):
    domains = set()

    for value in series.dropna():
        text = str(value).strip()

        # 支援：DM, AE / DM;AE / 換行
        parts = re.split(r"[,\n;]+", text)

        for part in parts:
            item = part.strip()
            if item:
                domains.add(item.upper())

    return domains


if uploaded_file is not None:
    try:
        # ✅ 讀全部 sheet 名稱
        xls = pd.ExcelFile(uploaded_file)
        all_sheets = xls.sheet_names

        # ✅ 檢查 SoA
        if "SoA" not in all_sheets:
            st.error("找不到 SoA 分頁")
        else:
            soa_df = pd.read_excel(uploaded_file, sheet_name="SoA", header=1)

            # ✅ !!! 關鍵：清理欄位名稱
            soa_df.columns = soa_df.columns.str.strip()

            # ✅ Debug：顯示欄位
            st.write("SoA 欄位名稱:", list(soa_df.columns))

            # ✅ 找 Form OID（容錯）
            form_oid_col = None

            for col in soa_df.columns:
                if "FORM" in col.upper() and "OID" in col.upper():
                    form_oid_col = col
                    break

            if form_oid_col is None:
                st.error("找不到 Form OID 欄位")
            else:
                st.success(f"使用欄位: {form_oid_col}")

                # ✅ 抓 domain
                valid_domains = extract_form_oids(soa_df[form_oid_col])

                st.write("SoA 定義的 Domain:", sorted(valid_domains))

                # ✅ Excel實際存在的 sheet（轉大寫）
                sheet_upper_map = {s.upper(): s for s in all_sheets}

                # ✅ 有交集的
                available_sheets = [
                    sheet_upper_map[d]
                    for d in valid_domains
                    if d in sheet_upper_map
                ]

                # ✅ SoA有但Excel沒有（很重要）
                missing_sheets = [
                    d for d in valid_domains
                    if d not in sheet_upper_map
                ]

                # ✅ 顯示 missing
                if missing_sheets:
                    st.warning(f"SoA 中有但 Excel 沒有的 Sheet: {missing_sheets}")

                # ✅ 顯示可選Sheet
                if not available_sheets:
                    st.error("沒有可顯示的 sheet")
                else:
                    st.success("已依 SoA 過濾 sheet")

                    selected_sheet = st.selectbox(
                        "請選擇 CRF Domain",
                        sorted(available_sheets)
                    )

                    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=3)

                    st.subheader(f"Sheet: {selected_sheet}")

                    st.write(f"資料筆數: {len(df)}")
                    st.write("欄位名稱:", list(df.columns))

                    st.dataframe(df)

    except Exception as e:
        st.error(f"發生錯誤: {e}")

