import streamlit as st
import pandas as pd
import re
from io import BytesIO
import pyreadstat

st.set_page_config(page_title="Auto SDTM SPEC", layout="wide")
st.title("Auto SDTM SPEC")

# =========================================================
# 工具函式（Step 1）
# =========================================================
def normalize_text(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x)).strip().upper()

def normalize_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df

def find_column(columns, keywords):
    for col in columns:
        upper = col.upper()
        if all(k in upper for k in keywords):
            return col
    return None

# =========================================================
# Header detect
# =========================================================
def detect_header(df, keywords, scan=20):
    for i in range(min(scan, len(df))):
        row = df.iloc[i].astype(str).str.upper()
        for cell in row:
            if all(k in cell for k in keywords):
                return i
    return None

# =========================================================
# parse SDTM target
# =========================================================
def parse_targets(val):
    if pd.isna(val):
        return []
    tokens = re.split(r"[;\n]+", str(val))
    results = []

    for t in tokens:
        t = t.strip()
        if "." in t:
            dom, rest = t.split(".", 1)
            var = rest.split("=")[0].strip().upper()
            val = ""
            if "=" in t:
                val = t.split("=")[1].strip("\"' ")
            results.append((dom.strip().upper(), var, val))
    return results

# =========================================================
# Step 1：Mapping
# =========================================================
def process_excel(file_bytes, manual_soa=None, manual_domain=None):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheets = xls.sheet_names

    # --- SoA ---
    raw_soa = pd.read_excel(BytesIO(file_bytes), "SoA", header=None)

    if manual_soa:
        header_row = manual_soa - 1
    else:
        header_row = detect_header(raw_soa, ["FORM", "OID"])

    soa = pd.read_excel(BytesIO(file_bytes), "SoA", header=header_row)
    soa = normalize_columns(soa)

    col = find_column(soa.columns, ["FORM", "OID"])
    domains = set(soa[col].dropna().astype(str).str.upper())

    # --- mapping ---
    records = []
    detail = []

    for s in sheets:
        if s.upper() not in domains:
            continue

        raw = pd.read_excel(BytesIO(file_bytes), s, header=None)

        if manual_domain:
            h = manual_domain - 1
        else:
            h = detect_header(raw, ["SDTM", "TARGET"])

        if h is None:
            continue

        df = pd.read_excel(BytesIO(file_bytes), s, header=h)
        df = normalize_columns(df)

        tgt_col = find_column(df.columns, ["SDTM", "TARGET"])
        src_col = df.columns[0]

        for _, r in df.iterrows():
            targets = parse_targets(r[tgt_col])

            for dom, var, assign in targets:
                records.append({"Dataset": dom, "Variable": var})

                detail.append({
                    "Dataset": dom,
                    "Variable": var,
                    "Source": s,
                    "CRF": r[src_col],
                    "Assign": assign
                })

    mapping_df = pd.DataFrame(records).drop_duplicates()
    detail_df = pd.DataFrame(detail)

    return mapping_df, detail_df

# =========================================================
# Step 2：SAS CONFIG
# =========================================================
def load_sas(var_file, ds_file):
    var_df = pd.DataFrame()
    ds_df = pd.DataFrame()

    if var_file:
        var_df, _ = pyreadstat.read_sas7bdat(var_file)
        var_df.columns = [c.strip() for c in var_df.columns]

    if ds_file:
        ds_df, _ = pyreadstat.read_sas7bdat(ds_file)
        ds_df.columns = [c.strip() for c in ds_df.columns]

    if "Dataset" in var_df:
        var_df["Dataset"] = var_df["Dataset"].str.upper()
        var_df["Variable"] = var_df["Variable"].str.upper()

    if "Dataset" in ds_df:
        ds_df["Dataset"] = ds_df["Dataset"].str.upper()

    return var_df, ds_df

# =========================================================
# build spec
# =========================================================
def build_var_spec(mapping, config):
    mapping_pairs = set(zip(mapping["Dataset"], mapping["Variable"]))

    crf = mapping.copy()

    if not config.empty:
        crf = crf.merge(config, on=["Dataset", "Variable"], how="left")

    # non-CRF
    if not config.empty:
        cfg_pairs = set(zip(config["Dataset"], config["Variable"]))
        non_pairs = cfg_pairs - mapping_pairs
        config["pair"] = list(zip(config["Dataset"], config["Variable"]))
        non_crf = config[config["pair"].isin(non_pairs)].drop(columns="pair")
    else:
        non_crf = pd.DataFrame()

    final = pd.concat([crf, non_crf]).drop_duplicates()

    return final.sort_values(["Dataset", "Variable"])

def build_ds_spec(var_spec, config):
    ds = pd.DataFrame({"Dataset": sorted(var_spec["Dataset"].unique())})

    if not config.empty:
        ds = ds.merge(config, on="Dataset", how="left")

    return ds

# =========================================================
# UI
# =========================================================
uploaded = st.file_uploader("Upload CRF Excel", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.read()

    # RESET session
    if "file" not in st.session_state or st.session_state["file"] != uploaded.name:
        st.session_state["file"] = uploaded.name
        st.session_state["run"] = False

    # =========================
    # Header override
    # =========================
    st.markdown("### Header Override")

    c1, c2 = st.columns(2)

    with c1:
        use_soa = st.checkbox("SoA header override")
        soa_header = st.number_input("Row", 1, 50, 2) if use_soa else None

    with c2:
        use_dom = st.checkbox("Domain header override")
        dom_header = st.number_input("Row ", 1, 50, 2) if use_dom else None

    # =========================
    # Step 1
    # =========================
    st.markdown("## Step 1｜CRF → SDTM Mapping")

    mapping_df, detail_df = process_excel(file_bytes, soa_header, dom_header)

    st.write("Mapping Summary")
    st.dataframe(mapping_df)

    st.write("Mapping Detail")
    st.dataframe(detail_df)

    st.session_state["mapping"] = mapping_df

    # =========================
    # Step 2 trigger
    # =========================
    if st.button("▶ 執行 Step 2"):
        st.session_state["run"] = True

    # =========================
    # Step 2
    # =========================
    if st.session_state["run"]:

        st.markdown("## Step 2｜SPEC Generator")

        st.markdown("### Upload SAS config")

        var_file = st.file_uploader("Variables (.sas7bdat)", type=["sas7bdat"])
        ds_file = st.file_uploader("Datasets (.sas7bdat)", type=["sas7bdat"])

        if var_file or ds_file:
            var_cfg, ds_cfg = load_sas(var_file, ds_file)
        else:
            var_cfg, ds_cfg = pd.DataFrame(), pd.DataFrame()

        var_spec = build_var_spec(mapping_df, var_cfg)
        st.markdown("### Variables SPEC")
        st.dataframe(var_spec)

        ds_spec = build_ds_spec(var_spec, ds_cfg)
        st.markdown("### Datasets SPEC")
        st.dataframe(ds_spec)

        # download
        out = BytesIO()
        with pd.ExcelWriter(out) as writer:
            var_spec.to_excel(writer, sheet_name="Variables", index=False)
            ds_spec.to_excel(writer, sheet_name="Datasets", index=False)

        st.download_button("Download SPEC", out.getvalue(), "SPEC.xlsx")
