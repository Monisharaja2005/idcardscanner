import cv2
import pytesseract
from PIL import Image
import pandas as pd
import re
import os
from datetime import datetime

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'

# Create directories for saving images
os.makedirs('captured_ids', exist_ok=True)
os.makedirs('processed_images', exist_ok=True)

# Initialize Excel file
excel_file = "id_data.xlsx"
if os.path.exists(excel_file):
    df = pd.read_excel(excel_file)
else:
    df = pd.DataFrame(columns=[
        "Timestamp", 
        "Name", 
        "ID Number", 
        "Department", 
        "Original Image Path",
        "Processed Image Path",
        "Raw Text"
    ])

def extract_id_info(text):
    """Extract name, ID number, and department from OCR text"""
    name = "Not Found"
    id_num = "Not Found"
    dept = "Not Found"
    
    # Try to extract name (adjust patterns based on your ID format)
    name_match = re.search(r'(?:Name|NAME)\s*[:]?\s*([A-Za-z ]+)', text)
    if name_match:
        name = name_match.group(1).strip()
    
    # Try to extract ID number
    id_match = re.search(r'(?:ID|Number|NO)\s*[:]?\s*([A-Za-z0-9-]+)', text)
    if id_match:
        id_num = id_match.group(1).strip()
    
    # Try to extract department
    dept_match = re.search(r'(?:Department|Dept|DEPT)\s*[:]?\s*([A-Za-z ]+)', text)
    if dept_match:
        dept = dept_match.group(1).strip()
    
    return name, id_num, dept

def process_image_for_ocr(img_path):
    """Process image to improve OCR accuracy and save processed version"""
    # Open and convert to grayscale
    img = Image.open(img_path).convert('L')
    
    # Enhance contrast and binarize
    img = img.point(lambda x: 0 if x < 128 else 255)  # Simple threshold
    
    # Save processed image
    processed_path = os.path.join('processed_images', 'processed_' + os.path.basename(img_path))
    img.save(processed_path)
    
    return processed_path, img

def capture_and_process_id():
    cap = cv2.VideoCapture(0)
    captured = False
    
    while True:
        ret, frame = cap.read()
        if not ret:
            print("Failed to capture frame")
            break
            
        # Display instructions
        cv2.putText(frame, "Align ID and press 'c' to capture", (10, 30), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)
        cv2.putText(frame, "Press 'q' to quit", (10, 60), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)
        
        # Show the frame
        cv2.imshow("ID Capture", frame)
        
        key = cv2.waitKey(1)
        
        if key & 0xFF == ord('q'):
            break
        elif key & 0xFF == ord('c'):
            # Save captured image with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            original_path = os.path.join('captured_ids', f'id_{timestamp}.jpg')
            cv2.imwrite(original_path, frame)
            print(f"ID captured and saved to {original_path}")
            captured = True
            break
    
    cap.release()
    cv2.destroyAllWindows()
    
    if captured:
        # Process the captured image
        processed_path, processed_img = process_image_for_ocr(original_path)
        print(f"Processed image saved to {processed_path}")
        
        # Perform OCR
        text = pytesseract.image_to_string(processed_img)
        print("\nRaw OCR Text:\n", text)
        
        # Extract information
        name, id_num, dept = extract_id_info(text)
        
        print("\nExtracted Information:")
        print(f"Name: {name}")
        print(f"ID Number: {id_num}")
        print(f"Department: {dept}")
        
        # Add to Excel
        global df
        new_entry = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Name": name,
            "ID Number": id_num,
            "Department": dept,
            "Original Image Path": original_path,
            "Processed Image Path": processed_path,
            "Raw Text": text[:1000]  # Store first 1000 chars of raw text
        }
        
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        df.to_excel(excel_file, index=False)
        print(f"\nAll data saved to {excel_file}")

        # Show the processed image
        processed_img.show()

if __name__ == "__main__":
    capture_and_process_id()
    print("\nProgram completed.")
