import os, shutil
from ultralytics import YOLO
import cv2

MODEL_PATH = "yolov8s.pt"  # Gunakan versi 's', 'm', 'l', atau 'x' untuk akurasi lebih tinggi
INPUT_FOLDER = "input_images"
OUTPUT_FOLDER = "output"

def detect_objects():
    model = YOLO(MODEL_PATH)
    
    if os.path.exists(OUTPUT_FOLDER):
        shutil.rmtree(OUTPUT_FOLDER)
    os.makedirs(OUTPUT_FOLDER)
        
    for image_name in os.listdir(INPUT_FOLDER):
        input_path = os.path.join(INPUT_FOLDER, image_name)
        output_path = os.path.join(OUTPUT_FOLDER, image_name)

        if not input_path.lower().endswith(('.png', '.jpg', '.jpeg')):
            continue
        
        results = model.predict(input_path, conf=0.25, imgsz=960, save=False)

        for result in results:
            annotated_image = result.plot()
            cv2.imwrite(output_path, annotated_image)
            print(f"Hasil Deteksi Disimpan ke: {output_path}")
            print(f"Jumlah Deteksi: {len(result.boxes)}")
            
if __name__ == "__main__":
    detect_objects()
