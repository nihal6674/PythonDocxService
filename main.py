from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import traceback
from fastapi.middleware.cors import CORSMiddleware
import subprocess
import tempfile
from pathlib import Path

from io import BytesIO
import boto3
import os
import qrcode
import re

from PIL import Image

from dotenv import load_dotenv

load_dotenv()
LIBREOFFICE_PATH = os.getenv(
    "LIBREOFFICE_PATH"
)

app = FastAPI()

# -----------------------------
# CORS CONFIG
# -----------------------------
cors_origins = os.getenv("CORS_ORIGINS", "")

allowed_origins = [o.strip() for o in cors_origins.split(",") if o.strip()]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins or [],  # empty = block all
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
INTERNAL_API_KEY = os.getenv("INTERNAL_API_KEY")

if not INTERNAL_API_KEY:
    raise RuntimeError("INTERNAL_API_KEY not set")

# -----------------------------
# R2 CONFIG
# -----------------------------
R2_ACCOUNT_ID = os.getenv("R2_ACCOUNT_ID")
R2_ACCESS_KEY_ID = os.getenv("R2_ACCESS_KEY_ID")
R2_SECRET_ACCESS_KEY = os.getenv("R2_SECRET_ACCESS_KEY")
R2_BUCKET_NAME = os.getenv("R2_BUCKET_NAME")

R2_ENDPOINT = f"https://{R2_ACCOUNT_ID}.r2.cloudflarestorage.com"
VERIFY_BASE_URL = os.getenv("VERIFY_BASE_URL")

if not VERIFY_BASE_URL:
    raise RuntimeError("VERIFY_BASE_URL not set")

s3 = boto3.client(
    "s3",
    endpoint_url=R2_ENDPOINT,
    aws_access_key_id=R2_ACCESS_KEY_ID,
    aws_secret_access_key=R2_SECRET_ACCESS_KEY,
    region_name="auto",
)

# -----------------------------
# HELPERS
# -----------------------------

def convert_docx_to_pdf(docx_bytes: bytes) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        docx_path = tmpdir / "input.docx"
        pdf_path = tmpdir / "input.pdf"

        docx_path.write_bytes(docx_bytes)

        subprocess.run(
            [
                LIBREOFFICE_PATH,
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                "--nodefault",
                "--nolockcheck",
                "--norestore",
                "--convert-to",
                "pdf",
                str(docx_path),
                "--outdir",
                str(tmpdir),
            ],
            check=True,
            capture_output=True,
            text=True,
        )

        if not pdf_path.exists():
            raise RuntimeError("PDF conversion failed")

        return pdf_path.read_bytes()

def download_from_r2(key: str) -> bytes:
    obj = s3.get_object(Bucket=R2_BUCKET_NAME, Key=key)
    return obj["Body"].read()

def upload_to_r2(key: str, data: bytes):
    s3.put_object(
        Bucket=R2_BUCKET_NAME,
        Key=key,
        Body=data,
        ContentType=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )

from datetime import datetime

def format_mmddyyyy(date_str: str) -> str:
    """
    Accepts:
    - YYYY-MM-DD
    - ISO datetime
    Returns:
    - MM/DD/YYYY
    """
    try:
        dt = datetime.fromisoformat(date_str)
        return dt.strftime("%m/%d/%Y")
    except Exception:
        return date_str  # fallback safely

# -----------------------------
# REQUEST SCHEMA
# -----------------------------
class GenerateDocxPayload(BaseModel):
    templateKey: str
    signatureKey: str
    outputKey: str
    data: dict

def safe_part(value: str) -> str:
    """
    Make filename-safe strings:
    - remove special chars
    - replace spaces with underscore
    """
    value = value.strip()
    value = re.sub(r"\s+", "_", value)
    value = re.sub(r"[^a-zA-Z0-9_]", "", value)
    return value

# -----------------------------
# ROUTES
# -----------------------------
@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/generate-docx")
def generate_docx(payload: GenerateDocxPayload, request: Request):
    api_key = request.headers.get("x-internal-api-key")

    if api_key != INTERNAL_API_KEY:
        raise HTTPException(status_code=403, detail="Forbidden")

    try:
        # 1️⃣ Download template
        template_bytes = download_from_r2(payload.templateKey)
        tpl = DocxTemplate(BytesIO(template_bytes))

        # 2️⃣ Generate QR
        cert_no = payload.data.get("certificate_number", "")
        if not cert_no:
            raise HTTPException(400, "certificate_number missing")

        verify_url = f"{VERIFY_BASE_URL.rstrip('/')}/{cert_no}"
        qr_img = qrcode.make(verify_url)
        qr_buf = BytesIO()
        qr_img.save(qr_buf, format="PNG")
        qr_buf.seek(0)


        # 3️⃣ Download instructor signature
        signature_bytes = download_from_r2(payload.signatureKey)

        # 🔽 DOWN-SCALE signature image safely
        img = Image.open(BytesIO(signature_bytes))
        img = img.convert("RGBA")   # normalize
        img.thumbnail((800, 300))   # max width x height

        sign_buf = BytesIO()
        img.save(sign_buf, format="PNG")
        sign_buf.seek(0)



        # 4️⃣ Context (TEXT + QR + SIGNATURE)
        context = {
            "first_name": payload.data.get("first_name", ""),
            "middle_name": payload.data.get("middle_name", ""),
            "last_name": payload.data.get("last_name", ""),
            "training_date": format_mmddyyyy(
                payload.data.get("training_date", "")
            ),
            "issue_date": format_mmddyyyy(
                payload.data.get("issue_date", "")
            ),
            "certificate_number": payload.data.get("certificate_number", ""),
            "instructor_name": payload.data.get("instructor_name", ""),
            "qr_code": InlineImage(tpl, qr_buf, width=Mm(30)),
            "instructor_signature": InlineImage(
                tpl, sign_buf, width=Mm(30)
            ),
        }


        tpl.render(context)

        # 5️⃣ Save DOCX to memory
        out = BytesIO()
        tpl.save(out)
        out.seek(0)

        cert_no = safe_part(payload.data.get("certificate_number", ""))
        first = safe_part(payload.data.get("first_name", ""))
        middle = safe_part(payload.data.get("middle_name", ""))
        last = safe_part(payload.data.get("last_name", ""))

        if not cert_no or not first or not last:
            raise HTTPException(400, "certificate_number, first_name and last_name are required")

        filename_parts = [cert_no, first]

        if middle:
            filename_parts.append(middle)

        filename_parts.append(last)

        filename = "_".join(filename_parts) + ".docx"

        # Optional: put into a folder
        output_key = f"certificates/{filename}"

        # Convert DOCX → PDF
        pdf_bytes = convert_docx_to_pdf(out.getvalue())

        pdf_key = output_key.replace(".docx", ".pdf")

        # Upload PDF
        s3.put_object(
            Bucket=R2_BUCKET_NAME,
            Key=pdf_key,
            Body=pdf_bytes,
            ContentType="application/pdf",
        )

        return {"key": pdf_key}



    except Exception as e:
        traceback.print_exc() 
        raise HTTPException(status_code=500, detail=str(e))
