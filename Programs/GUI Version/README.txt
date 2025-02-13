Python version: Python 3.13.1

For DETS Serial Number Maker:
You should put in Excel folder student data in Students Data.xlsx file [Name, Acadmic Number]

For DETS Correction System:
Just follow the instructions

Required packages:
OpenCV: pip install opencv-python-headless opencv-python
NumPy: pip install numpy
openpyxl: pip install openpyxl
pytesseract: pip install pytesseract
Tesseract OCR: https://github.com/tesseract-ocr/tesseract/blob/main/README.md
shutil: pip install pytest-shutil
Pillow: pip install Pillow
customtkinter: pip install customtkinter
sys: pip install os-sys
threading: pip install threaded

Important Note:
Don't forget to change line 20 on DETS Correction System (pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe)
if the path of tesseract.exe is different