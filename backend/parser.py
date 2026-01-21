import re
import pandas as pd
from docx import Document

def normalize(text: str) -> str:
    return text.replace("\xa0", " ").strip()

def parse_cell(text: str):
    text = normalize(text)

    serial = None
    name = None
    enrollment = None
    contact = None
    address = None

    # SERIAL
    m = re.search(r"\b(\d{5})\b", text)
    if m:
        serial = m.group(1)

    # NAME (between serial and D/)
    m = re.search(r"\d{5}\s+([A-Z\s]+?)\s+D/", text)
    if m:
        name = m.group(1).strip()

    # ENROLLMENT
    m = re.search(r"(D/\d+/\d{4})", text)
    if m:
        enrollment = m.group(1)

    # CONTACT
    m = re.search(r"Contact:\s*([0-9]{10})", text)
    if m:
        contact = m.group(1)

    # ADDRESS (everything after Address:)
    m = re.search(r"Address:\s*([\s\S]+)", text)
    if m:
        address = " ".join(m.group(1).split())

    return {
        "Serial": serial,
        "Name": name,
        "Enrollment": enrollment,
        "Contact": contact,
        "Address": address
    }

def parse_docx_to_excel(docx_path, excel_path):
    doc = Document(docx_path)
    records = []

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if not text:
                    continue

                parsed = parse_cell(text)

                # record is valid only if name + enrollment exist
                if parsed["Name"] and parsed["Enrollment"]:
                    records.append(parsed)

    df = pd.DataFrame(records)
    df.to_excel(excel_path, index=False)
