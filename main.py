# main.py - ไฟล์หลักสำหรับรัน API

from flask import Flask, request, jsonify, send_file
import os
import tempfile
from werkzeug.utils import secure_filename
from extract_images import extract_images_from_xlsx
import base64
import io

app = Flask(__name__)

@app.route('/')
def home():
    return "XLSX Image Extractor API - ส่ง POST request ไปที่ /extract พร้อมแนบไฟล์ .xlsx ในฟิลด์ 'file'"

@app.route('/extract', methods=['POST'])
def extract():
    if 'file' not in request.files:
        return jsonify({"error": "ไม่พบไฟล์ในคำขอ"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "ไม่ได้เลือกไฟล์"}), 400
    
    if not file.filename.endswith('.xlsx'):
        return jsonify({"error": "รองรับเฉพาะไฟล์ .xlsx เท่านั้น"}), 400
    
    # สร้างโฟลเดอร์ชั่วคราวสำหรับเก็บไฟล์
    temp_dir = tempfile.mkdtemp()
    filename = secure_filename(file.filename)
    file_path = os.path.join(temp_dir, filename)
    file.save(file_path)
    
    try:
        # ดึงรูปภาพออกจากไฟล์ Excel
        images_data = extract_images_from_xlsx(file_path)
        
        # แปลงรูปภาพเป็น base64 สำหรับส่งกลับเป็น JSON
        result = []
        for i, image_data in enumerate(images_data):
            img_io = image_data['image_data']
            img_format = image_data.get('format', 'PNG')
            img_base64 = base64.b64encode(img_io.getvalue()).decode('utf-8')
            
            result.append({
                "index": i,
                "sheet_name": image_data.get('sheet_name', 'unknown'),
                "cell_address": image_data.get('cell_address', 'unknown'),
                "format": img_format,
                "base64": img_base64
            })
        
        return jsonify({
            "status": "success",
            "count": len(result),
            "images": result
        })
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
    finally:
        # ลบไฟล์ชั่วคราว
        try:
            os.remove(file_path)
            os.rmdir(temp_dir)
        except:
            pass

if __name__ == '__main__':
    # ตั้งค่าที่จำเป็นสำหรับ Railway
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)