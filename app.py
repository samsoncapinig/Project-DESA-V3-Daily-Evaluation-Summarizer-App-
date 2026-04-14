import streamlit as st
import pandas as pd
import numpy as np
import re

from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER

# =============================
# PAGE CONFIG
# =============================
st.set_page_config(page_title="Evaluation Dashboard", layout="wide")
st.title("📊 Evaluation Dashboard (Fixed Version)")

# =============================
# LOAD FILE
# =============================
def load_any_file(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except:
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file)

# =============================
# HELPERS
# =============================
def compute_avg_for_columns(df, cols):
    if not cols:
        return np.nan
    sub = df[cols].copy()
    arr = pd.to_numeric(sub.to_numpy().ravel(), errors="coerce")
    return np.nanmean(arr)

def generate_recommendations(weaknesses_df):
    if weaknesses_df is None or weaknesses_df.empty:
        return ["Continue maintaining current performance standards."]

    recs = []
    for _, row in weaknesses_df.head(3).iterrows():
        recs.append(f"Improve {row['Category']} by enhancing delivery and resources.")
    return recs

# =============================
# UPLOAD
# =============================
uploaded_files = st.file_uploader(
    "Upload CSV or Excel Files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True
)

all_category_combined = []
overall_all = []

if uploaded_files:

    for file in uploaded_files:

        st.divider()
        st.subheader(file.name)

        df = load_any_file(file)

        if df is None:
            continue

        st.success("Loaded")

        df_num = df.apply(pd.to_numeric, errors="coerce")

        numeric_cols = df_num.columns[df_num.notna().any()].tolist()

        rating_cols = [
            c for c in numeric_cols
            if "id" not in c.lower() and "response" not in c.lower()
        ]

        if rating_cols:
            overall_avg = compute_avg_for_columns(df, rating_cols)
            overall_all.append(overall_avg)

            st.metric("Overall Rating", round(overall_avg, 2))

            category_map = {}
            for col in rating_cols:
                category_map[col] = col.split("_")[0]

            cat_data = []
            for col in rating_cols:
                cat_data.append({
                    "Category": category_map[col],
                    "Average": pd.to_numeric(df[col], errors="coerce").mean()
                })

            category_df = pd.DataFrame(cat_data)

            category_avg = category_df.groupby("Category", as_index=False).mean()
            category_avg["File"] = file.name

            st.dataframe(category_avg)
            st.bar_chart(category_avg.set_index("Category"))

            all_category_combined.append(category_avg)

# =============================
# SUMMARY
# =============================
if all_category_combined:

    combined = pd.concat(all_category_combined, ignore_index=True)
    combined = combined.dropna(subset=["Average"])

    summary = combined.groupby("Category", as_index=False)["Average"].mean()

    strengths = summary.sort_values("Average", ascending=False).head(3)
    weaknesses = summary.sort_values("Average", ascending=True).head(3)

    st.subheader("Top Strengths")
    st.dataframe(strengths)

    st.subheader("Top Weaknesses")
    st.dataframe(weaknesses)

    st.subheader("Recommendations")
    for r in generate_recommendations(weaknesses):
        st.write("•", r)
