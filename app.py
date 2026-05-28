

import streamlit as st
import pandas as pd

st.title("Excel Sheet Viewer")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    
    st.write("Sheet 列表:", sheets)
    
    selected = st.selectbox("選擇 Sheet", sheets)
    
    df = pd.read_excel(uploaded_file, sheet_name=selected)
    st.dataframe(df)
