import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Auto SDTM SPEC", layout="wide")
st.title("Auto SDTM SPEC")


# =========================================================
# 工具
# =========================================================
def normalize_text(x):
    if pd.isna(x):
        return ""
    x = str(x)
    x = re.sub(r"\s+", " ", x)
    return x.strip().upper()


def normalize_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_column(columns, keywords):
    for col in columns:
        if all(k in col.upper() for k in keywords):
            return col
    return None


# =========================================================
# Header detect（修正版）
# =========================================================
def detect_header(df, keywords, scan=30):
    for i in range(min(scan, len(df))):
        row = df.iloc[i].values

        for cell in row:
            cell_str = str(cell).upper()   # ✅ 防 crash

            if all(k.upper() in cell_str for k in keywords):
                return i

    return None


# =========================================================
# SDTM parse
# =========================================================
def parse_targets(val):
    if pd.isna(val):
        return []

    tokens = re.split(r"[;\n]+", str(val))
    out = []

    for t in tokens:
        t = t.strip()

        if "." in t:
            dom, rest = t.split(".", 1)
            var = rest.split("=")[0].strip().upper()

            assign = ""
            if "=" in t:
                assign = t.split("=")[1].strip("\"' ")

            out.append((dom.strip().upper(), var, assign))

    return out


# =========================================================
# Step 1 主邏輯
# =========================================================
def process_excel(file_bytes, soa_header=None, domain_header=None):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheets = xls.sheet_names

    # =====================
    # SoA
    # =====================
    raw_soa = pd.read_excel(BytesIO(file_bytes), "SoA", header=None)

    if soa_header:
        h = soa_header - 1
    else:
        h = detect_header(raw_soa, ["FORM", "OID"])

    if h is None:
        raise ValueError("SoA header 無法辨識")

    soa = pd.read_excel(BytesIO(file_bytes), "SoA", header=h)
    soa = normalize_columns(soa)

    form_col = find_column(soa.columns, ["FORM", "OID"])
    if form_col is None:
        raise ValueError("找不到 Form OID 欄位")

    domains = set(
        soa[form_col]
        .dropna()
        .astype(str)
        .str.upper()
    )

    # =====================
    # Mapping
    # =====================
    mapping = []
    detail = []
    sheet_errors = []

    for s in sheets:
        if s.upper() not in domains:
            continue

        raw = pd.read_excel(BytesIO(file_bytes), s, header=None)

        if domain_header:
            dh = domain_header - 1
        else:
            dh = detect_header(raw, ["SDTM", "TARGET"])

        if dh is None:
            sheet_errors.append(s)
            continue

        df = pd.read_excel(BytesIO(file_bytes), s, header=dh)
        df = normalize_columns(df)

        tgt_col = find_column(df.columns, ["SDTM", "TARGET"])
        src_col = df.columns[0]

        for idx, r in df.iterrows():
            targets = parse_targets(r[tgt_col])

            for dom, var, assign in targets:
                mapping.append({"SDTM Domain": dom, "SDTM Variable": var})

                detail.append({
                    "Source Sheet": s,
                    "CRF Variable": r[src_col],
                    "SDTM Domain": dom,
                    "SDTM Variable": var,
                    "Assign": assign
                })

    mapping_df = pd.DataFrame(mapping).drop_duplicates()
    detail_df = pd.DataFrame(detail)

    return mapping_df, detail_df, sheet_errors


# =========================================================
# UI
# =========================================================
uploaded = st.file_uploader("Upload CRF Mapping Excel", type=["xlsx"])

if uploaded:

    file_bytes = uploaded.getvalue()

    # reset step2 when file changes
    if "file" not in st.session_state or st.session_state["file"] != uploaded.name:
        st.session_state["file"] = uploaded.name
        st.session_state["run"] = False

    # =====================
    # Header override
    # =====================
    st.markdown("### Header Override")

    c1, c2 = st.columns(2)

    with c1:
        use_soa = st.checkbox("SoA header override")
        soa_header = None
        if use_soa:
            soa_header = st.number_input("Row", 1, 50, 2)

    with c2:
        use_dom = st.checkbox("Domain header override")
        dom_header = None
        if use_dom:
            dom_header = st.number_input("Row ", 1, 50, 2)

    # =====================
    # Step 1
    # =====================
    st.markdown("## Step 1｜CRF → SDTM Mapping")

    try:
        mapping_df, detail_df, sheet_errors = process_excel(
            file_bytes,
            soa_header,
            dom_header
        )

        st.session_state["mapping"] = mapping_df

        # summary
        st.markdown("### Summary")
        if mapping_df.empty:
            st.warning("No SDTM mapping found")
        else:
            summary = (
                mapping_df.groupby("SDTM Domain")["SDTM Variable"]
                .apply(lambda x: sorted(set(x)))
                .reset_index()
            )
            summary["Count"] = summary["SDTM Variable"].apply(len)
            summary["Variables"] = summary["SDTM Variable"].apply(lambda x: "; ".join(x))

            st.dataframe(summary[["SDTM Domain", "Count", "Variables"]])

        # detail
        st.markdown("### Detail")
        st.dataframe(detail_df)

        # error
        if sheet_errors:
            st.warning(f"Header 偵測失敗：{sheet_errors}")

    except Exception as e:
        st.error(f"Step 1 Error: {e}")
        st.stop()

    # =====================
    # Step 2 trigger
    # =====================
    if st.button("▶ 執行 Step 2"):
        st.session_state["run"] = True

    # =====================
    # Step 2（先只做入口）
    # =====================
    if st.session_state.get("run", False):
        st.markdown("## Step 2｜SPEC Generator")

        st.info("✅ 已成功進入 Step 2（下一步我們再接 config）")

        st.write("目前 mapping 筆數:", len(st.session_state.get("mapping", [])))
