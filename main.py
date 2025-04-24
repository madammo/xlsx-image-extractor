from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os, shutil, uuid
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# === ตั้งค่าการเชื่อม Google Drive ===
gauth = GoogleAuth()
gauth.LoadCredentialsFile("credentials.json")
if gauth.credentials is None:
    gauth.LocalWebserverAuth()
elif gauth.access_token_expired:
    gauth.Refresh()
else:
    gauth.Authorize()
gauth.SaveCredentialsFile("credentials.json")
drive = GoogleDrive(gauth)

# === สร้าง API ===
app = FastAPI()

@app.get("/")
def root():
    return {"message": "FastAPI for image upload to Google Drive is running 🎉"}

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
            image_data = img.image if hasattr(img, 'image') else img
            file_name = f"slip_{uuid.uuid4().hex}.png"
            local_path = os.path.join(temp_dir, file_name)
            image_data.save(local_path)

            # Upload ไป Google Drive
            gfile = drive.CreateFile({
                'title': file_name,
                'parents': [{'id': '1g1UedBOYIcIiqOn2MJxgWnQz4Gk2isTY'}]
            })
            gfile.SetContentFile(local_path)
            gfile.Upload()
            gfile.InsertPermission({'type': 'anyone', 'value': 'anyone', 'role': 'reader'})
            uploaded_links.append(gfile['alternateLink'])

        return {"uploaded_slips": uploaded_links}

    except Exception as e:
        return JSONResponse(content={"error": str(e)}, status_code=500)
