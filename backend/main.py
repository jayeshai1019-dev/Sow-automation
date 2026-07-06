from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from dotenv import load_dotenv
import os

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
    # Toggles for optional sections
    include_landing_zone:      bool = False
    include_control_tower:     bool = False
    include_landing_zone_arch: bool = False
    include_paloalto:          bool = False
    include_mgn_migration:     bool = False

import traceback

@app.post("/generate-sow")
async def generate_sow(request: SOWRequest):
    try:
        output_path = generate_sow_document(request.dict())
        filename = f"SOW_{request.customer_name.replace(' ', '_')}.docx"
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
