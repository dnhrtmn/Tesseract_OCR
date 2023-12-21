import cv2
import pytesseract
import tkinter as tk
from tkinter import filedialog
import openpyxl

pytesseract.pytesseract.tesseract_cmd = '/usr/local/bin/tesseract'  

def extract_text_from_image(image, region=None):
    if region is not None:
        x, y, w, h = region
        image = image[y:y + h, x:x + w]

    text = pytesseract.image_to_string(image)
    return text

def select_region(image_path):
    image = cv2.imread(image_path)

    cv2.imshow('Select Region', image)
    rect_box = cv2.selectROI('Select Region', image, fromCenter=False, showCrosshair=True)
    cv2.destroyAllWindows()

    return rect_box

def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Choose a File", filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.ico")])

    if not file_path:
        print("File selection canceled.")
        return

    region = select_region(file_path)

    if region == (0, 0, 0, 0):
        print("Region selection canceled.")
        return

    if file_path.lower().endswith(('.png', '.jpg', '.jpeg')):
        image = cv2.imread(file_path)
        extracted_text = extract_text_from_image(image, region)
    else:
        print("Unsupported file format.")
        return

    excel_file = 'output.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet['A1'] = 'Extracted Text'
    sheet['B1'] = extracted_text

    workbook.save(excel_file)
    print(f"Extracted text has been written to {excel_file}.")

if __name__ == "__main__":
    main()