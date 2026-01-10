from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from io import BytesIO
import boto3
import os
import qrcode
from PIL import Image

from dotenv import load_dotenv

load_dotenv()

app = FastAPI()

# -----------------------------
# R2 CONFIG
# -----------------------------
R2_ACCOUNT_ID = os.getenv("R2_ACCOUNT_ID")
R2_ACCESS_KEY_ID = os.getenv("R2_ACCESS_KEY_ID")
R2_SECRET_ACCESS_KEY = os.getenv("R2_SECRET_ACCESS_KEY")
R2_BUCKET_NAME = os.getenv("R2_BUCKET_NAME")

R2_ENDPOINT = f"https://{R2_ACCOUNT_ID}.r2.cloudflarestorage.com"

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

# -----------------------------
# REQUEST SCHEMA
# -----------------------------
class GenerateDocxPayload(BaseModel):
    templateKey: str
    signatureKey: str
    outputKey: str
    data: dict

# -----------------------------
# ROUTES
# -----------------------------
@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/generate-docx")
def generate_docx(payload: GenerateDocxPayload):
    try:
        # 1️⃣ Download template
        template_bytes = download_from_r2(payload.templateKey)
        tpl = DocxTemplate(BytesIO(template_bytes))

        # 2️⃣ Generate QR
        cert_no = payload.data.get("certificate_number", "")
        if not cert_no:
            raise HTTPException(400, "certificate_number missing")

        qr_img = qrcode.make(cert_no)
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
            "training_date": payload.data.get("training_date", ""),
            "issue_date": payload.data.get("issue_date", ""),
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

        # 6️⃣ Upload DOCX to R2
        upload_to_r2(payload.outputKey, out.getvalue())

        return {"key": payload.outputKey}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
