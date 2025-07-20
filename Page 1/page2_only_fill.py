import pandas as pd
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfObject, PdfName
import os

TEMPLATE_PATH = 'Second Summer template_Ver2.pdf'
CSV_PATH = 'P9 test 1.csv'
OUTPUT_DIR = 'output_page2'

os.makedirs(OUTPUT_DIR, exist_ok=True)
df = pd.read_csv(CSV_PATH)
filename_counts = {}

def fill_pdf(data, output_path):
    template_pdf = PdfReader(TEMPLATE_PATH)

    # === FORCE Times-Bold Font registration ===
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
    annotations = page['/Annots']

    if annotations:
        for annotation in annotations:
            if annotation['/Subtype'] == '/Widget' and annotation.get('/T'):
                key = annotation['/T'][1:-1]  # Strip parentheses

                # Determine value
                if key == "Student":
                    value = str(data.get("Student", ""))
                elif key in data:
                    value = str(data[key])
                else:
                    continue

                # === FORCE FONT PER FIELD ===
                annotation.update(PdfDict(
                    V=PdfObject(f'({value})'),
                    Ff=1,
                    DA="/T1 12 Tf 0 g",  # Force Times-Bold 12pt font
                    MK=PdfDict(BC='[0 0 0]'),  # Optional: Border color black
                    FT=PdfName('Tx'),
                    DR=PdfDict(Font=PdfDict(
                        T1=PdfDict(
                            Type=PdfName.Font,
                            Subtype=PdfName.Type1,
                            BaseFont=PdfName('Times-Bold')
                        )
                    ))
                ))

                # === Overlay for flattening just "Student" field visually ===
                if key == "Student":
                    contents = page.Contents.stream if '/Contents' in page else ''
                    overlay = f"\nBT /T1 12 Tf 1 0 0 1 125 715 Tm ({value}) Tj ET\n"
                    page.Contents.stream = PdfObject(f"{contents}{overlay}")

    # === Ensure visibility in all PDF viewers ===
    acroform.update(PdfDict(NeedAppearances=PdfObject('true')))

    PdfWriter(output_path, trailer=template_pdf).write()

# === Generate PDFs ===
for index, row in df.iterrows():
    student_name = row.get("Student", f"student_{index}")
    base_name = student_name.replace(" ", "_").replace(",", "")
    count = filename_counts.get(base_name, 0)
    filename_counts[base_name] = count + 1

    filename = f"{base_name}{'' if count == 0 else f'_{count}'}.pdf"
    output_path = os.path.join(OUTPUT_DIR, filename)

    print(f"Creating PDF for: {student_name}")
    fill_pdf(row.to_dict(), output_path)

print("✅ Page 2 PDFs created in 'output_page2'")
