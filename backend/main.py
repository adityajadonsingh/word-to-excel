from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from parser import parse_cell
from docx import Document
import json
import os
import uuid

app = FastAPI()

# CORS (dev-safe)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

SESSION_FILE = "session_records.json"
OUTPUT_EXCEL = "final_output.xlsx"

# init session
if not os.path.exists(SESSION_FILE):
    with open(SESSION_FILE, "w") as f:
        json.dump([], f)

# -------------------------------
# UPLOAD WORD (append records)
# -------------------------------
@app.post("/upload-docx")
async def upload_docx(file: UploadFile = File(...)):
    temp_path = f"temp_{uuid.uuid4()}.docx"

    with open(temp_path, "wb") as f:
        f.write(await file.read())

    doc = Document(temp_path)

    new_records = []

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parsed = parse_cell(cell.text)
                if parsed["Name"] and parsed["Enrollment"]:
                    new_records.append(parsed)

    # append to session
    with open(SESSION_FILE, "r") as f:
        existing = json.load(f)

    seen = {(r["Serial"], r["Enrollment"]) for r in existing}

    for r in new_records:
        key = (r["Serial"], r["Enrollment"])
        if key not in seen:
            existing.append(r)
            seen.add(key)

    with open(SESSION_FILE, "w") as f:
        json.dump(existing, f, indent=2)

    os.remove(temp_path)

    return {
        "added": len(new_records),
        "total": len(existing)
    }

# -------------------------------
# DOWNLOAD EXCEL
# -------------------------------
@app.get("/download-excel")
def download_excel():
    import pandas as pd

    with open(SESSION_FILE, "r") as f:
        data = json.load(f)

    if not data:
        return JSONResponse(
            {"error": "No data uploaded"},
            status_code=400
        )

    df = pd.DataFrame(data)
    df.to_excel(OUTPUT_EXCEL, index=False)

    return FileResponse(
        OUTPUT_EXCEL,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="final_output.xlsx"
    )

# -------------------------------
# RESET SESSION 
# -------------------------------
@app.post("/reset-session")
def reset_session():
    with open(SESSION_FILE, "w") as f:
        json.dump([], f)
    return {"status": "session cleared"}
