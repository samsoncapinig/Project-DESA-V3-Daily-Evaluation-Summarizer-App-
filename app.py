import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import tempfile
from typing import Dict, List, Tuple, Optional

# Plotting
import plotly.express as px

# PPTX
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

# =============================
# PAGE CONFIG
# =============================
st.set_page_config(page_title="Evaluation Dashboard", layout="wide")

st.title("📊 Evaluation Dashboard (Combined Version)")
st.caption("Auto-detects CSV / Excel files | Generates summaries, insights, and reports")

# =============================
# UNIVERSAL FILE LOADER
# =============================
def load_any_file(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file)
        except Exception as e:
            st.error(f"❌ Unsupported file: {e}")
            return None

# =============================
# HELPERS
# =============================
def coerce_numeric(df):
    return df.apply(pd.to_numeric, errors="coerce")

def compute_avg(df, cols):
    if not cols:
        return np.nan
    return coerce_numeric(df[cols]).stack().mean()

# =============================
# INSIGHTS DETECTION
# =============================
INSIGHT_REGEX = re.compile(r"insight", re.IGNORECASE)

def find_insight_cols(df):
    return [c for c in df.columns if INSIGHT_REGEX.search(str(c))]

def extract_insights(df):
    cols = find_insight_cols(df)
    positives, improvements = [], []

    for col in cols:
        for val in df[col].dropna():
            text = str(val).lower()

            if any(w in text for w in ["good", "excellent", "great", "helpful"]):
                positives.append(val)
            elif any(w in text for w in ["improve", "should", "need", "lack"]):
                improvements.append(val)

    return positives, improvements

# =============================
# FILE UPLOAD
# =============================
uploaded_files = st.file_uploader(
    "Upload CSV or Excel Files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True
)

# =============================
# MAIN PROCESS
# =============================
if uploaded_files:
    
    all_category = []
    all_session = []
    all_category_combined = []
    
    insights_all = {"positive": [], "improve": []}

    for file in uploaded_files:

        st.divider()
        st.subheader(f"📄 {file.name}")

        df = load_any_file(file)

        if df is None:
            continue

        st.success("File loaded successfully")

        # =============================
        # NUMERIC DETECTION
        # =============================
        # Try to coerce all columns first (more flexible)
        df_numeric = df.apply(pd.to_numeric, errors="coerce")

        numeric_cols = df_numeric.columns[df_numeric.notna().any()].tolist()

        rating_cols = [
            col for col in numeric_cols
            if not any(x in col.lower() for x in ["id", "response"])
        ]

        if rating_cols:
            overall_avg = compute_avg(df, rating_cols)

            if not np.isnan(overall_avg):
                st.metric("Overall Rating", round(overall_avg, 2))
            else:
                st.warning("⚠ Unable to compute overall rating.")
            st.metric("Overall Rating", round(overall_avg, 2))

# =============================
# CATEGORY SUMMARY (ADAPTIVE)
# =============================
if rating_cols:

    category_map = {}

    for col in rating_cols:
        if "->" in str(col):
            category = col.split("->")[0].strip()
        else:
            # fallback: group by prefix before underscore or use full name
            category = str(col).split("_")[0]

        category_map[col] = category

    category_data = []

    for col in rating_cols:
        if col in df.columns:
            avg_val = pd.to_numeric(df[col], errors="coerce").mean()
            category_data.append({
                "Category": category_map[col],
                "Average": avg_val
            })

    if category_data:
        category_df = pd.DataFrame(category_data)

        category_avg = (
            category_df
            .groupby("Category", as_index=False)
            .mean()
        )

        category_avg["File"] = file.name

        st.dataframe(category_avg)

        if not category_avg.empty:
            st.bar_chart(category_avg.set_index("Category"))

        all_category.append(category_avg)
        all_category_combined.append(category_avg)
    else:
        st.warning("⚠ No valid category data found.")

else:
    st.warning("⚠ No rating columns detected.")
        

    # =============================
    # TOP 3 STRENGTHS & WEAKNESSES
    # =============================
    st.divider()
    st.subheader("🏆 Top 3 Strengths & Weaknesses")

    if all_category_combined:

        combined_df = pd.concat(all_category_combined, ignore_index=True)

        # Remove invalid values
        combined_df = combined_df.dropna(subset=["Average"])

        if not combined_df.empty:

            # Compute overall average per category across files
            summary = (
                combined_df
                .groupby("Category", as_index=False)["Average"]
                .mean()
            )

            # Sort
            top_strengths = summary.sort_values(by="Average", ascending=False).head(3)
            top_weaknesses = summary.sort_values(by="Average", ascending=True).head(3)

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### ✅ Top Strengths")
                for _, row in top_strengths.iterrows():
                st.write(f"**{row['Category']}** — {row['Average']:.2f}")

            with col2:
                st.markdown("### ⚠️ Top Weaknesses")
                for _, row in top_weaknesses.iterrows():
                    st.write(f"**{row['Category']}** — {row['Average']:.2f}")

        else:
            st.warning("⚠ No valid category data available.")

    else:
        st.info("No category data to summarize.")

    def generate_narrative(overall_avg, strengths_df, weaknesses_df, insights):
    lines = []

    # =============================
    # OVERALL PERFORMANCE
    # =============================
    if overall_avg is not None and not np.isnan(overall_avg):
        if overall_avg >= 4.5:
            remark = "an excellent level of satisfaction"
        elif overall_avg >= 4.0:
            remark = "a very satisfactory level of performance"
        elif overall_avg >= 3.0:
            remark = "a satisfactory level, with room for improvement"
        else:
            remark = "a low level of satisfaction, indicating the need for significant improvements"

        lines.append(
            f"Overall, the evaluation results indicate {remark}, with a mean rating of {overall_avg:.2f}."
        )

    # =============================
    # STRENGTHS
    # =============================
    if strengths_df is not None and not strengths_df.empty:
        top_items = ", ".join(
            [f"{row['Category']} ({row['Average']:.2f})" for _, row in strengths_df.iterrows()]
        )

        lines.append(
            f"The highest-rated areas were {top_items}, suggesting that these aspects were particularly effective and well-received by participants."
        )

    # =============================
    # WEAKNESSES
    # =============================
    if weaknesses_df is not None and not weaknesses_df.empty:
        low_items = ", ".join(
            [f"{row['Category']} ({row['Average']:.2f})" for _, row in weaknesses_df.iterrows()]
        )

        lines.append(
            f"On the other hand, the lowest-rated areas included {low_items}, indicating opportunities for improvement in these aspects."
        )

    # =============================
    # QUALITATIVE INSIGHTS
    # =============================
    pos_count = len(insights.get("positive", []))
    imp_count = len(insights.get("improve", []))

    if pos_count > 0:
        lines.append(
            f"Qualitative feedback further highlighted several strengths, with participants expressing positive remarks on the program’s effectiveness and delivery."
        )

    if imp_count > 0:
        lines.append(
            f"However, some respondents also provided suggestions for improvement, particularly in areas related to delivery, resources, and overall experience."
        )

    # =============================
    # CONCLUSION
    # =============================
    lines.append(
        "Overall, the results suggest that while the program is performing well, continuous enhancements in identified areas will further improve participant satisfaction and outcomes."
    )

    return " ".join(lines)

    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.platypus import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
from io import BytesIO

def line_field(width=400):
    return "_" * int(width / 5)


def generate_exact_form5(exec_summary, overall_avg, weaknesses):

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=30,
        bottomMargin=30
    )

    styles = getSampleStyleSheet()

    normal = ParagraphStyle(name="normal", fontSize=10)
    center = ParagraphStyle(name="center", alignment=TA_CENTER, fontSize=10)
    bold_center = ParagraphStyle(name="bold_center", alignment=TA_CENTER, fontSize=11, spaceAfter=6)

    elements = []

    # =============================
    # HEADER (EXACT POSITIONING)
    # =============================
    elements.append(Paragraph("RV-QAME     TOOL-05", normal))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("<b>TITLE</b>", center))
    elements.append(Paragraph("<b>Overall Monitoring & Evaluation Results Form</b>", bold_center))

    elements.append(Paragraph("Page 1 of 1", center))
    elements.append(Spacer(1, 12))

    # =============================
    # INSTRUCTIONS
    # =============================
    elements.append(Paragraph(
        "Instructions: This monitoring tool will be accomplished by the onsite monitor/s. "
        "This will form part of the QAME reports to be submitted to the program owner/PMT "
        "a week after the conduct of the training program.",
        normal
    ))

    elements.append(Spacer(1, 6))

    elements.append(Paragraph(
        "Data Privacy Statement: All the data to be generated will be treated with utmost confidentiality "
        "and shall be governed by Republic Act 10173, otherwise known as the Data Privacy Act of 2012.",
        normal
    ))

    elements.append(Spacer(1, 12))

    # =============================
    # FORM FIELDS
    # =============================
    fields = [
        "Title of Training Program",
        "Date and Venue",
        "Learning Service Provider/Division",
        "Learning Areas"
    ]

    for f in fields:
        elements.append(Paragraph(f"{f}", normal))
        elements.append(Paragraph(line_field(450), normal))
        elements.append(Spacer(1, 8))

    # =============================
    # PARTICIPANTS
    # =============================
    elements.append(Paragraph("No. & Description of Participants", normal))
    elements.append(Paragraph(line_field(450), normal))

    elements.append(Spacer(1, 6))

    elements.append(Paragraph("Teaching", normal))
    elements.append(Paragraph(line_field(200), normal))

    elements.append(Paragraph("Non-Teaching", normal))
    elements.append(Paragraph(line_field(200), normal))

    elements.append(Paragraph("Teaching Related", normal))
    elements.append(Paragraph(line_field(200), normal))

    elements.append(Spacer(1, 12))

    # =============================
    # RESULTS TABLE (MATCHED)
    # =============================
    results_data = [
        ["Result of Daily Online Evaluation", ""],
        ["Result of End-of-Program Evaluation", ""],
        ["Overall Result", f"{overall_avg:.2f}" if overall_avg else ""],
    ]

    table = Table(results_data, colWidths=[3.5*inch, 2.5*inch])
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE")
    ]))

    elements.append(table)
    elements.append(Spacer(1, 12))

    # =============================
    # ANALYSIS (MULTI-LINE EXACT)
    # =============================
    elements.append(Paragraph("<b>Analysis:</b>", normal))
    elements.append(Spacer(1, 6))

    for _ in range(5):
        elements.append(Paragraph(line_field(450), normal))

    elements.append(Spacer(1, 12))

    # =============================
    # RECOMMENDATIONS (MULTI-LINE)
    # =============================
    elements.append(Paragraph("<b>Recommendations:</b>", normal))
    elements.append(Spacer(1, 6))

    recs = generate_recommendations(weaknesses)

    for r in recs:
        elements.append(Paragraph(f"• {r}", normal))

    for _ in range(3):
        elements.append(Paragraph(line_field(450), normal))

    elements.append(Spacer(1, 24))

    # =============================
    # SIGNATURES (MATCHED LAYOUT)
    # =============================
    signature_table = Table([
        ["Prepared by:", "", "Noted:"],
        ["", "", ""],
        ["____________________", "", "____________________"],
        ["Name of QAME Monitor", "", "SGOD Chief"]
    ], colWidths=[2.5*inch, 1*inch, 2.5*inch])

    signature_table.setStyle(TableStyle([
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
    ]))

    elements.append(signature_table)

    # =============================
    # BUILD
    # =============================
    doc.build(elements)
    buffer.seek(0)

    return buffer
