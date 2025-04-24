from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os, shutil, uuid

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SERVICE_ACCOUNT_FILE = '/etc/secrets/n8ncheckslip-7a8a20abc6b5.json'  # <== ใส่ไฟล์ JSON ที่ได้จากขั้นตอน 1
FOLDER_ID = '1g1UedBOYIcIiqOn2MJxgWnQz4Gk2isTY'  # <== Folder ID ที่คุณส่งมา

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=['https://www.googleapis.com/auth/drive']
)
drive_service = build('drive', 'v3', credentials=credentials)

app = FastAPI()

@app.get("/")
def root():
    return {"message": "API with Service Account is running 🎉"}

@app.post("/extract-images")
async def extract_images(file: UploadFile = File(...)):
    try:
        temp_dir = "temp_uploads"
        os.makedirs(temp_dir, exist_ok=True)

        temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4()}_{file.filename}")
        with open(temp_file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        wb = load_workbook(temp_file_path)
        ws = wb.active

        uploaded_links = []

        for i, img in enumerate(ws._images):
            img_data = img.image if hasattr(img, 'image') else img
            file_name = f"slip_{uuid.uuid4().hex}.png"
            local_path = os.path.join(temp_dir, file_name)
            img_data.save(local_path)

            media = MediaFileUpload(local_path, mimetype='image/png')
            file_metadata = {
                'name': file_name,
                'parents': [FOLDER_ID]
            }
            uploaded_file = drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id,webViewLink'
            ).execute()

            uploaded_links.append(uploaded_file['webViewLink'])

        return {"uploaded_slips": uploaded_links}

    except Exception as e:
        return JSONResponse(content={"error": str(e)}, status_code=500)
