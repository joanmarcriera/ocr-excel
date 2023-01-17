#!/usr/bin/env python3

import pytesseract
import cv2
import pandas as pd
from PIL import Image
from openpyxl import Workbook

# Import other OCR libraries
try:
    import pyocr
    from pyocr import builders
    pyocr_available = True
except ImportError:
    pyocr_available = False

try:
    import textract
    textract_available = True
except ImportError:
    textract_available = False

photo="captured_image.jpg"



# Open the laptop camera
camera = cv2.VideoCapture(0)

# Capture an image
while True:
    ret, frame = camera.read()
    cv2.imshow('Press SPACE to click...', frame)
    key = cv2.waitKey(1)
    if key == ord(' '):
        break

cv2.imwrite(photo, frame)
camera.release()

# Perform OCR using pytesseract
image = Image.open("captured_image.jpg")
text_pytesseract = pytesseract.image_to_string(image)

print ("1"+text_pytesseract)
# Create a DataFrame from the OCR output
data = [line.split() for line in text_pytesseract.split("\n") if line.strip()]
df_pytesseract = pd.DataFrame(data)

# Write the DataFrame to an Excel file
writer = pd.ExcelWriter("pytesseract_output.xlsx", engine='openpyxl')
df_pytesseract.to_excel(writer,sheet_name='Sheet1', index=False, header=False)
writer.save()
writer.close()

# Perform OCR using pyocr (if available)
if pyocr_available:
    tools = pyocr.get_available_tools()
    if tools:
        tool = tools[0]
        text_pyocr = tool.image_to_string(
            Image.open("captured_image.jpg"),
            lang="eng",
            builder=builders.TextBuilder()
        )
        print ("2"+text_pyocr)
        # Create a DataFrame from the OCR output
        data = [line.split() for line in text_pyocr.split("\n") if line.strip()]
        df_pyocr = pd.DataFrame(data)

        # Write the DataFrame to an Excel file
        
        writer = pd.ExcelWriter("pyocr_output.xlsx", engine='openpyxl')
        
        df_pyocr.to_excel(writer,sheet_name='Sheet1', index=False, header=False)
        writer.save()
        writer.close()
    else:
        print("No OCR tool found")


# Perform OCR using textract (if available)
if textract_available:
    text_textract = textract.process("captured_image.jpg", method='tesseract', language='eng')
    text_textract = text_textract.decode("utf-8")
    print ("3"+text_textract)
    # Create a DataFrame from the OCR output
    data = [line.split() for line in text_textract.split("\n") if line.strip()]
    df_textract = pd.DataFrame(data)

    # Write the DataFrame to an Excel file
    
    writer = pd.ExcelWriter("textract_output.xlsx", engine='openpyxl')
    
    df_textract.to_excel(writer,sheet_name='Sheet1', index=False, header=False)
    writer.save()
    writer.close()

