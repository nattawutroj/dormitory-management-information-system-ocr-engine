from flask import Flask, request, jsonify
from flask_cors import CORS
from io import BytesIO
from PIL import Image
from ultralytics import YOLO
import cv2
import numpy as np
import easyocr
import re

app = Flask(__name__)
CORS(app)  # Allow CORS for localhost

# Load YOLO model and OCR reader
model = YOLO("./src/model/best.pt")
reader = easyocr.Reader(['en'])  # Using English language for OCR

@app.route('/upload', methods=['POST'])
def upload_image():
    # Check if an image file is included in the request
    if 'image' not in request.files:
        return jsonify({"error": "No file provided"}), 400
    
    file = request.files['image']
    
    # Validate that the file is an image
    if file and file.content_type.startswith('image'):
        try:
            # Open and resize the image if necessary
            image = Image.open(file)
            max_height = 640
            if image.height > max_height:
                height_percent = max_height / float(image.height)
                new_width = int(float(image.width) * height_percent)
                image = image.resize((new_width, max_height), Image.LANCZOS)
            
            # Convert image to numpy array for YOLO processing
            img_np = np.array(image.convert("RGB"))
            
            # Run YOLO prediction with confidence threshold
            results = model.predict(img_np, conf=0.5)
            
            # Prepare a blank background for annotated output
            bg_width, bg_height = 800, 600
            background = np.zeros((bg_height, bg_width, 3), dtype=np.uint8)
            
            # Draw bounding boxes and labels on the background image
            for result in results:
                boxes = result.boxes
                for box in boxes:
                    x1, y1, x2, y2 = map(int, box.xyxy[0])
                    label = result.names[int(box.cls)]
                    
                    # Only label non-panel detections
                    if label != "panel":
                        label_position = (x1, y1 - 10 if y1 - 10 > 10 else y1 + 10)
                        cv2.putText(background, label, label_position, cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
            
            # Save the annotated image temporarily
            labeled_img_path = "./src/tem/labeled_image.jpg"
            cv2.imwrite(labeled_img_path, background)
            
            # OCR processing on the labeled image
            ocr_image = Image.open(labeled_img_path)
            ocr_result = reader.readtext(np.array(ocr_image), detail=0)
            
            # Extract only numeric characters
            numbers_only = " ".join(re.findall(r'\d+', " ".join(ocr_result)))
            final = numbers_only.replace(" ", "")
            
            # Return OCR result as JSON
            return jsonify({"ocr_result": final}), 200
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    else:
        return jsonify({"error": "File is not an image"}), 400

@app.route('/upload', methods=['GET'])
def hello_world():
    return jsonify("hello world /upload"), 200

if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0", port=5000)