#!/usr/bin/env python3
import sys
from PyPDF2 import PdfReader
import pandas as pd

def extract_data_from_pdf(pdf_path):
    """
    Extracts key fields from the RCS PDF based on label matches.
    Returns a dict mapping field keys to extracted text.
    """
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        page_text = page.extract_text() or ""
        text += page_text + "\n"

    data = {}

    # ——— EXAMPLE EXTRACTIONS ———
    # These look for lines beginning with your PDF labels. Add/update as needed.

    for line in text.splitlines():
        line = line.strip()
        if line.startswith("As Of:"):
            # splits on colon, takes right side
            data['effective_date'] = line.split(":", 1)[1].strip()
        # elif line.startswith("Prepared For:"):
        #     data['prepared_for'] = line.split(":", 1)[1].strip()
        # elif line.startswith("Market Area:"):
        #     data['market_area'] = line.split(":", 1)[1].strip()
        # … you can add more patterns here …

    return data

def fill_excel(template_path, output_path, data):
    """
    Loads the blank Excel and populates the
    “9-5-2 Detailed Screening” tab with values from data dict,
    then saves to output_path.
    """
    # read only that sheet into DataFrame
    df = pd.read_excel(
        template_path,
        sheet_name="9-5-2 Detailed Screening",
        engine="openpyxl"
    )

    # ——— EXAMPLE MAPPING ———
    # Find the row where the “Field” column equals your PDF label,
    # then write the extracted value into the “Notes” column.
    mask = df['Field'] == "Part A - Question 1"
    if mask.any():
        df.loc[mask, 'Notes'] = data.get('effective_date', '')

    # … add one block like the above for each PDF field you’re extracting …

    # write back into the same sheet
    with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        df.to_excel(writer,
            sheet_name="9-5-2 Detailed Screening",
            index=False
        )

def main():
    if len(sys.argv) != 4:
        print("Usage: python3 rcs_app.py <input_pdf> <template_xlsx> <output_xlsx>")
        sys.exit(1)

    pdf_path      = sys.argv[1]
    template_path = sys.argv[2]
    output_path   = sys.argv[3]

    data = extract_data_from_pdf(pdf_path)
    fill_excel(template_path, output_path, data)
    print(f"✅ Completed: '{output_path}' created.")

if __name__ == "__main__":
    main()
