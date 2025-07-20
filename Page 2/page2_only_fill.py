import pandas as pd
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfObject, PdfName
import os
import re

# === File paths ===
TEMPLATE_PATH = 'Second Summer template_Ver2.pdf'
CSV_PATH = 'p9 test 1.csv'
OUTPUT_DIR = 'output_page2'

os.makedirs(OUTPUT_DIR, exist_ok=True)
df = pd.read_csv(CSV_PATH)

# === Track filename counts ===
filename_counts = {}

# === Utility to clean non-Latin-1 characters ===
def clean_text(value):
    if not isinstance(value, str):
        value = str(value)
    replacements = {
        '\u2018': "'",  # left single quote
        '\u2019': "'",  # right single quote
        '\u201C': '"',  # left double quote
        '\u201D': '"',  # right double quote
        '\u2013': '-',  # en dash
        '\u2014': '-',  # em dash
        '\u2026': '...',  # ellipsis
        '\xa0': ' ',    # non-breaking space
    }
    for bad, good in replacements.items():
        value = value.replace(bad, good)
    # Remove other characters outside Latin-1
    return value.encode('latin-1', errors='ignore').decode('latin-1')

# === Fill only Page 2 (Interim Report) ===
def fill_pdf(data, output_path):
    template_pdf = PdfReader(TEMPLATE_PATH)

    # === Register Times-Bold font as /T1 globally ===
    if '/AcroForm' not in template_pdf.Root:
        template_pdf.Root.AcroForm = PdfDict()

    acroform = template_pdf.Root.AcroForm

    if not acroform.get('DR'):
        acroform.DR = PdfDict()
    if not acroform.DR.get('Font'):
        acroform.DR.Font = PdfDict()

    acroform.DR.Font.update({
        PdfName('T1'): PdfDict(
            Type=PdfName.Font,
            Subtype=PdfName.Type1,
            BaseFont=PdfName('Times-Bold')
        )
    })

    acroform.update(PdfDict(DA="/T1 12 Tf 0 g"))

    page = template_pdf.pages[0]
    annotations = page.get('/Annots')

    if annotations:
        for annotation in annotations:
            if annotation['/Subtype'] == '/Widget' and annotation.get('/T'):
                key = annotation['/T'][1:-1]
                if key in data:
                    value = clean_text(data[key])  # CLEANED HERE
                    # === Update each field to use Times-Bold explicitly ===
                    annotation.update(PdfDict(
                        V=PdfObject(f'({value})'),
                        Ff=4096,
                        DA="/T1 12 Tf 0 g",
                        DR=PdfDict(Font=PdfDict(
                            T1=PdfDict(
                                Type=PdfName.Font,
                                Subtype=PdfName.Type1,
                                BaseFont=PdfName('Times-Bold')
                            )
                        ))
                    ))

    # === Ensure text appears properly in viewers ===
    acroform.update(PdfDict(NeedAppearances=PdfObject('true')))

    PdfWriter(output_path, trailer=template_pdf).write()

# === Generate PDFs ===
for index, row in df.iterrows():
    base_name = clean_text(row.get("Student_1", f"student_{index}")).replace(" ", "_").replace(",", "")
    count = filename_counts.get(base_name, 0)
    filename_counts[base_name] = count + 1

    filename = f"{base_name}{count}_page2.pdf"
    output_path = os.path.join(OUTPUT_DIR, filename)

    fill_pdf(row.to_dict(), output_path)

print("✅ Page 2 PDFs created in 'output_page2'")
