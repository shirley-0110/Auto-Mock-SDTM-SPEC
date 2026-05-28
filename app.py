import streamlit as st
import pandas as pd
import re
from io import BytesIO


def normalize_text(x):
    """清理字串，方便比對欄位名稱"""
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
    檢查一整列是否包含指定關鍵字組
    keyword_groups 是 list of list，例如：
    [["FORM", "OID"]] 或 [["SDTM", "TARGET"]]
    代表只要某個儲存格同時包含該組關鍵字，就算找到
    """
    cells = [normalize_text(v) for v in row_values]

    for cell in cells:
        for group in keyword_groups:
            if all(k in cell for k in group):
                return True
    return False


def detect_header_row(file_bytes, sheet_name, keyword_groups, max_scan_rows=30):
    """
    掃描前幾列，找出哪一列包含目標欄位名稱
    回傳 0-based row index；若找不到回傳 None
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
    先偵測 header row，再重新讀 sheet
    manual_header_row_excel:
        若有值，代表使用者手動指定 Excel 第幾列為 header（1-based）
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

    # 回傳 Excel 人類看到的列號（1-based）
    return df, header_row_zero_based + 1
