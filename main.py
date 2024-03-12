import datetime
from pathlib import Path
from docxtpl import DocxTemplate
import pandas as pd
from tqdm import tqdm

base_dir = Path(__file__).parent
word_template_path = base_dir / "kby_python.docx"
excel_path = base_dir / "kby_anaconda.xlsx"
output_dir = base_dir / "Output"

output_dir.mkdir(exist_ok=True)

df = pd.read_excel(excel_path, sheet_name="Sheet1")
count = 0
print(df)
for record in df.to_dict(orient="records"):
    count += 1
    context = {
        "B": record["RSK Name"],
        "C": record["Beneficiary name"],
        "D": record["Kby Mesurement"],
        "E": record['Village'],
        "F": record['Gram panchaythi name'],
        "G": record["SUBSIDY"],
        "H": record["SURVEY NUMBER"],
        "I": record["DATE"],
        "J": record["Category"],
        "K": record["AADHAR NUMBER"],
        "L": record["AREA"],
        "M": record["Phone number"],
        "N": record["FID NUMBER"],
        "O": record["FULLRATE"],
        "P": record["FARMER SHARE"],
        "Q": record["LOGITTUDE (N)"],
        "R": record["Lattitud(E)"],
        "S": record["SANCTION ORDER"]
    }
    print(context)
    doc = DocxTemplate(word_template_path)
    doc.render(context)
    doc.save(output_dir / f"{record['FID NUMBER']}-{record['Beneficiary name']}.docx")
print(f"Total Records Generated is {count}")
