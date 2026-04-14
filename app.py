import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Evaluation Dashboard FIXED", layout="wide")
st.title("📊 Evaluation Dashboard (FIXED VERSION v2)")

# =============================
# FILE LOADER
# =============================
def load_any_file(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except:
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file)

# =============================
# SAFE AVG (NO STACK)
# =============================
def compute_avg_for_columns(df, cols):
    if not cols:
        return np.nan

    cols = [c for c in cols if c in df.columns]
    if not cols:
        return np.nan

    sub = df.loc[:, cols].copy()

    # remove duplicate columns (IMPORTANT FIX)
    sub = sub.loc[:, ~sub.columns.duplicated()]

    arr = pd.to_numeric(sub.to_numpy().ravel(), errors="coerce")
    return np.nanmean(arr)

# =============================
# CATEGORY SUMMARY SAFE
# =============================
def summarize_categories(df, rating_cols):
    category_map = {}

    for col in rating_cols:
        if "->" in str(col):
            category = str(col).split("->")[0].strip()
        else:
            category = str(col).split("_")[0]
        category_map[col] = category

    rows = []
    for col in rating_cols:
        if col in df.columns:
            val = pd.to_numeric(df[col], errors="coerce").mean()
            rows.append({
                "Category": category_map[col],
                "Average": val
            })

    if not rows:
        return pd.DataFrame()

    cat_df = pd.DataFrame(rows)

    # SAFE grouping (NO stack anywhere)
    return cat_df.groupby("Category", as_index=False).mean()

# =============================
# UPLOAD
# =============================
uploaded_files = st.file_uploader(
    "Upload CSV / Excel",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True
)

all_cat = []

if uploaded_files:

    for f in uploaded_files:

        st.divider()
        st.subheader(f.name)

        df = load_any_file(f)

        if df is None:
            continue

        st.success("Loaded successfully")

        df_num = df.apply(pd.to_numeric, errors="coerce")

        numeric_cols = df_num.columns[df_num.notna().any()].tolist()

        rating_cols = [
            c for c in numeric_cols
            if "id" not in str(c).lower() and "response" not in str(c).lower()
        ]

        if not rating_cols:
            st.warning("No rating columns found")
            continue

        overall = compute_avg_for_columns(df, rating_cols)
        st.metric("Overall Rating", round(overall, 2) if not np.isnan(overall) else 0)

        cat_df = summarize_categories(df, rating_cols)

        if not cat_df.empty:
            st.dataframe(cat_df)
            st.bar_chart(cat_df.set_index("Category"))

            cat_df["File"] = f.name
            all_cat.append(cat_df)

# =============================
# CROSS FILE SUMMARY
# =============================
if all_cat:

    merged = pd.concat(all_cat, ignore_index=True)
    merged = merged.dropna(subset=["Average"])

    summary = merged.groupby("Category", as_index=False)["Average"].mean()

    st.divider()
    st.subheader("📌 Cross-File Summary")

    st.dataframe(summary)

    st.subheader("🏆 Strengths")
    st.dataframe(summary.sort_values("Average", ascending=False).head(3))

    st.subheader("⚠ Weaknesses")
    st.dataframe(summary.sort_values("Average").head(3))
