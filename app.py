import streamlit as st
import pandas as pd
import re
from io import BytesIO
import pyreadstat

st.set_page_config(page_title="Auto SDTM SPEC", layout="wide")
st.title("Auto SDTM SPEC")


# =========================================================
# 基本工具
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
# SAS CONFIG LOADER
# =========================================================
def load_sas_config(variable_file, dataset_file):
    var_df = pd.DataFrame()
    ds_df = pd.DataFrame()

    if variable_file:
        var_df, _ = pyreadstat.read_sas7bdat(variable_file)
        var_df = normalize_columns(var_df)

    if dataset_file:
        ds_df, _ = pyreadstat.read_sas7bdat(dataset_file)
        ds_df = normalize_columns(ds_df)

    # normalize key columns
    for df in [var_df, ds_df]:
        if not df.empty:
            if "Dataset" in df.columns:
                df["Dataset"] = df["Dataset"].astype(str).str.upper()
            if "Variable" in df.columns:
                df["Variable"] = df["Variable"].astype(str).str.upper()

    return var_df, ds_df


# =========================================================
# Mapping logic（保留簡化版）
# =========================================================
def parse_sdtm_targets(value):
    results = []
    if pd.isna(value):
        return results

    tokens = re.split(r"[;\n]+", str(value))

    for token in tokens:
        token = token.strip()
        if "." in token:
            parts = token.split(".")
            if len(parts) >= 2:
                dom = parts[0].strip().upper()
                var = parts[1].split("=")[0].strip().upper()

                assign = ""
                if "=" in token:
                    assign = token.split("=")[1].strip("\"' ")

                results.append((dom, var, assign))

    return results


def build_mapping(df, sheet_name):
    col = find_column(df.columns, ["SDTM", "TARGET"])

    records = []

    for _, row in df.iterrows():
        targets = parse_sdtm_targets(row[col])

        for dom, var, assign in targets:
            records.append({
                "Dataset": dom,
                "Variable": var,
                "Assign": assign,
                "Source": sheet_name
            })

    return pd.DataFrame(records)


# =========================================================
# SPEC BUILDER (用 config)
# =========================================================
def build_variables_spec(mapping_df, config_var_df):

    mapping_pairs = set(zip(mapping_df["Dataset"], mapping_df["Variable"]))

    # CRF variables
    crf_df = mapping_df.copy()
    crf_df["Origin"] = crf_df["Assign"].apply(
        lambda x: "Assigned" if str(x).strip() else "CRF"
    )

    crf_df["Label"] = ""
    crf_df["Type"] = ""

    if not config_var_df.empty:
        merged = crf_df.merge(
            config_var_df,
            on=["Dataset", "Variable"],
            how="left",
            suffixes=("", "_CFG")
        )

        for col in ["Label", "Type", "Origin"]:
            cfg_col = f"{col}_CFG"
            if cfg_col in merged.columns:
                merged[col] = merged[col].replace("", merged[cfg_col])

        crf_df = merged

    # non-CRF variables
    if not config_var_df.empty:
        config_pairs = set(zip(config_var_df["Dataset"], config_var_df["Variable"]))
        non_crf_pairs = config_pairs - mapping_pairs

        non_crf_df = config_var_df.copy()
        non_crf_df["pair"] = list(zip(non_crf_df["Dataset"], non_crf_df["Variable"]))
        non_crf_df = non_crf_df[non_crf_df["pair"].isin(non_crf_pairs)]

        non_crf_df = non_crf_df.drop(columns="pair")
    else:
        non_crf_df = pd.DataFrame()

    final_df = pd.concat([crf_df, non_crf_df], ignore_index=True)

    final_df = final_df.drop_duplicates()
    final_df = final_df.sort_values(by=["Dataset", "Variable"])

    return final_df


def build_dataset_spec(var_spec_df, config_ds_df):
    datasets = sorted(var_spec_df["Dataset"].unique())

    df = pd.DataFrame({
        "Dataset": datasets
    })

    if not config_ds_df.empty:
        df = df.merge(config_ds_df, on="Dataset", how="left")

    return df


# =========================================================
# UI
# =========================================================
uploaded_file = st.file_uploader("上傳 CRF Mapping Excel", type=["xlsx"])

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names

    # STEP 1
    st.markdown("## Step 1｜CRF → SDTM Mapping")

    all_records = []

    for sheet in sheets:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        df = normalize_columns(df)

        col = find_column(df.columns, ["SDTM", "TARGET"])
        if col:
            tmp = build_mapping(df, sheet)
            all_records.append(tmp)

    if all_records:
        mapping_df = pd.concat(all_records)
    else:
        mapping_df = pd.DataFrame(columns=["Dataset", "Variable", "Assign", "Source"])

    st.dataframe(mapping_df)

    # STEP 2 BUTTON
    if st.button("▶ 執行 Step 2：SPEC Generator"):
        st.session_state["run"] = True

    if st.session_state.get("run"):

        st.markdown("## Step 2｜SPEC Generator")

        # CONFIG UPLOAD
        st.markdown("### 上傳 SAS config")

        var_file = st.file_uploader("Variables config (.sas7bdat)", type=["sas7bdat"])
        ds_file = st.file_uploader("Datasets config (.sas7bdat)", type=["sas7bdat"])

        config_var_df = pd.DataFrame()
        config_ds_df = pd.DataFrame()

        if var_file or ds_file:
            config_var_df, config_ds_df = load_sas_config(var_file, ds_file)

        # VARIABLES
        st.markdown("### Variables SPEC")

        var_spec = build_variables_spec(mapping_df, config_var_df)
        st.dataframe(var_spec)

        # DATASETS
        st.markdown("### Datasets SPEC")

        ds_spec = build_dataset_spec(var_spec, config_ds_df)
        st.dataframe(ds_spec)

        # DOWNLOAD
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            var_spec.to_excel(writer, sheet_name="Variables", index=False)
            ds_spec.to_excel(writer, sheet_name="Datasets", index=False)

        st.download_button(
            "下載 SPEC Excel",
            output.getvalue(),
            "SDTM_SPEC.xlsx"
        )
