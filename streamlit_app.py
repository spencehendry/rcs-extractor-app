import os
import io
import pandas as pd
import fitz  # PyMuPDF
import streamlit as st

st.title("RCS Detailed Screening Extractor")
st.markdown("Upload the blank RCS checklist and the RCS PDF to auto-fill key fields.")

# File upload widgets
template = st.file_uploader("📄 Upload Excel Checklist Template", type=["xlsx"])
pdf = st.file_uploader("📘 Upload RCS Appraisal PDF", type=["pdf"])

# Define prompts for each field
FIELD_PROMPTS = {
    "Part A - Question 1": "Provide the date and details of the inspection of the subject property as described in the RCS PDF.",
    "Part A - Question 2": "Describe any data collection issues noted in the RCS PDF.",
}

# Function to extract raw PDF text
def extract_text(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return text

# Dummy field extractor — this will be replaced with OpenAI API logic
def extract_field(text, prompt):
    return f"[Extracted answer for]: {prompt[:40]}..."  # Stub for now

# Run on button click
if st.button("Run Extraction"):
    if not template or not pdf:
        st.error("Please upload both the Excel and PDF files.")
    else:
        raw_text = extract_text(pdf)
        df = pd.read_excel(template, sheet_name="9-5-2 Detailed Screening", engine="openpyxl")

        for field, prompt in FIELD_PROMPTS.items():
            result = extract_field(raw_text, prompt)
            mask = df["Field"] == field
            df.loc[mask, "Notes"] = result

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="9-5-2 Detailed Screening", index=False)

        st.success("✅ Extraction complete!")
        st.download_button(
            label="📥 Download Populated Checklist",
            data=output.getvalue(),
            file_name="populated_detailed_screening.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
