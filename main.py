from flask import Flask, request, jsonify, send_file
from io import BytesIO
from PIL import Image
from ultralytics import YOLO
import cv2
import numpy as np
import easyocr
import re

app = Flask(__name__)

model = YOLO("./model/best.pt")
reader = easyocr.Reader(['en'])  # ใช้ English language model

@app.route('/upload', methods=['POST'])
def upload_image():
    
    if 'image' not in request.files:
        return {"error": "No file provided"}, 400
    
    file = request.files['image']
    
    if file and file.content_type.startswith('image'):
        image = Image.open(file)
        
        max_height = 640
        if image.height > max_height:
            height_percent = max_height / float(image.height)
            new_width = int((float(image.width) * float(height_percent)))
            image = image.resize((new_width, max_height), Image.LANCZOS)
        
        img_np = np.array(image.convert("RGB"))
        
        results = model.predict(img_np, conf=0.5)
        
        bg_width, bg_height = 800, 600
        background = np.zeros((bg_height, bg_width, 3), dtype=np.uint8)
        
        for result in results:
            boxes = result.boxes
            for box in boxes:
                x1, y1, x2, y2 = map(int, box.xyxy[0])
                label = result.names[int(box.cls)]
                
                if label != "panel":
                    label_position = (x1, y1 - 10 if y1 - 10 > 10 else y1 + 10)
                    cv2.putText(background, label, label_position, cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
        
        labeled_img_path = "./tem/labeled_image.jpg"
        cv2.imwrite(labeled_img_path, background)
        
        # OCR processing with easyocr
        ocr_image = Image.open(labeled_img_path)
        ocr_result = reader.readtext(np.array(ocr_image), detail=0)
        
        # Filter only digits using regular expressions
        numbers_only = " ".join(re.findall(r'\d+', " ".join(ocr_result)))
        
        final = numbers_only.replace(" ","")
        
        return jsonify({"ocr_result": final}), 200
    else:
        return {"error": "File is not an image"}, 400

if __name__ == '__main__':
    app.run(debug=True)