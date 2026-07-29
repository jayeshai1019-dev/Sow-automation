from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from dotenv import load_dotenv
import os
import tempfile

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

from sow_generator import generate_sow_document
from url_extractor import extract_company_info

app = FastAPI(title="SOW Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

class SOWRequest(BaseModel):
    customer_name: str
    mom_text: str
    company_url: Optional[str] = ""
    project_type: Optional[str] = "POC"   # "POC" or "Production"
    # Toggles for optional sections
    include_landing_zone:      bool = False
    include_control_tower:     bool = False
    include_landing_zone_arch: bool = False
    include_paloalto:          bool = False
    include_mgn_migration:     bool = False
    include_post_deployment:   bool = False
    include_testing_monitoring: bool = False
    include_monitoring:        bool = False
    include_dr:                bool = False

import traceback

@app.post("/generate-sow")
async def generate_sow(
    customer_name: str = Form(...),
    mom_text: str = Form(...),
    company_url: str = Form(""),
    project_type: str = Form("POC"),
    doc_date: str = Form(""),
    submitted_by: str = Form(""),
    include_landing_zone: str = Form("false"),
    include_control_tower: str = Form("false"),
    include_landing_zone_arch: str = Form("false"),
    include_paloalto: str = Form("false"),
    include_mgn_migration: str = Form("false"),
    include_post_deployment: str = Form("false"),
    include_testing_monitoring: str = Form("false"),
    include_monitoring: str = Form("false"),
    include_dr: str = Form("false"),
    client_logo: Optional[UploadFile] = File(None),
):
    try:
        # Parse string booleans from FormData
        def to_bool(val):
            return val.lower() in ('true', '1', 'yes')

        # Save uploaded logo to a temp file if provided
        logo_path = None
        if client_logo and client_logo.filename:
            logo_ext = os.path.splitext(client_logo.filename)[1] or ".png"
            logo_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=logo_ext)
            logo_tmp.write(await client_logo.read())
            logo_tmp.close()
            logo_path = logo_tmp.name

        data = {
            "customer_name": customer_name,
            "mom_text": mom_text,
            "company_url": company_url,
            "project_type": project_type,
            "doc_date": doc_date,
            "submitted_by": submitted_by,
            "include_landing_zone": to_bool(include_landing_zone),
            "include_control_tower": to_bool(include_control_tower),
            "include_landing_zone_arch": to_bool(include_landing_zone_arch),
            "include_paloalto": to_bool(include_paloalto),
            "include_mgn_migration": to_bool(include_mgn_migration),
            "include_post_deployment": to_bool(include_post_deployment),
            "include_testing_monitoring": to_bool(include_testing_monitoring),
            "include_monitoring": to_bool(include_monitoring),
            "include_dr": to_bool(include_dr),
            "client_logo_path": logo_path,
        }

        output_path = generate_sow_document(data)
        filename = f"SOW_{customer_name.replace(' ', '_')}.docx"

        # Cleanup logo temp file after generation
        if logo_path and os.path.exists(logo_path):
            os.unlink(logo_path)

        return FileResponse(
            path=output_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=filename,
        )
    except Exception as e:
        tb = traceback.format_exc()
        print(f"\n[ERROR] /generate-sow failed:\n{tb}")
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {str(e)}")

@app.get("/health")
def health():
    return {"status": "ok"}
