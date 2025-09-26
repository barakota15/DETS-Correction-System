# DETS Correction System
![Python Version](https://img.shields.io/badge/Python-3.13.1-blue) [![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

The DETS Correction System is an automated tool designed to correct exams by processing scanned images of answer sheets, extracting student responses, comparing them with the correct answers, and generating detailed Excel reports.

---

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [Installation](#installation)
4. [Usage](#usage)
5. [Source Code](#source-code)
6. [Folder Structure](#folder-structure)
8. [License](#license)
9. [Acknowledgments](#acknowledgments)

---

## Overview
The DETS Correction System automates the process of grading multiple-choice exams. It uses Optical Character Recognition (OCR) to extract student responses from scanned images, compares them with the correct answers, and generates Excel files containing the results. The system also supports serial number extraction and encoding for unique test identification.

![DETS Correction System](https://github.com/user-attachments/assets/0021f8fd-afc3-474d-9a20-064ee340e7c0)

---

## Features
- **Automated Grading**: Automatically grades exams based on predefined answer keys.
- **Serial Number Handling**: Extracts and encodes serial numbers from exam sheets.
- **Excel Reports**: Generates detailed Excel reports with color-coded correctness indicators.
- **User-Friendly GUI**: Provides a simple graphical user interface for ease of use.
- **Support for Multiple Templates**: Supports exams with 10, 30, or 60 questions.

---

## Installation

### Prerequisites
- Python 3.13.1
- Tesseract OCR: Install from [Sourceforge](https://sourceforge.net/projects/tesseract-ocr-alt/files/).
- Remeber to install it in `C:\Program Files\Tesseract-OCR\` path.

### Install Required Packages
Run the following command to install the required Python packages:

```bash
pip install -r requirements.txt
```

---

## Usage

### For DETS Serial Number Maker
```
Programs/
└── Serial Number Maker/
    ├── DETS Serial Number Maker.py # Python Code That Make Serial Numbers for Students
    └── Students Data.xlsx # Excel File That Should Contain Students Data
```

1. Place the student data in the `Students Data.xlsx` file inside the `Programs/Serial Number Maker`
2. Run the python code `DETS Serial Number Maker.py`.
3. Reopen the `Students Data.xlsx` file.

### For DETS Correction System

1. Run the program by python from the `Programs/GUI Version/GUI DETS Correction System.py`
2. Follow the on-screen instructions:
  - Choose the answer sheet template (10, 30, or 60)
  - Choose the folder containing the scanned exam images.
  - Select the Excel file containing student data.
  - Specify the correct answers for the exam.
3. Wait for the system to process the exams and generate the results.
4. Comfirm the results and save the excel file.

---

## Source Code

- For non-graphical code you can find it here: `Programs/DETS Correction System/DETS Correction System.py`
- For graphical code you can find it here: `Programs/GUI Version/GUI DETS Correction System.py`

---

## Folder Structure
```
DETS-Correction-System/
├── Exams/ # Examples for full real exams to test by the program
│   ├── 10 tests/
│   ├── 30 tests/
│   └── 60 tests/
├── Programs/
│   ├── DETS Correction System/ # The non-graphical program
│   |   ├── Data/
│   |   └── DETS Correction System.py
│   ├── GUI Version/ # The Graphical Program
│   |   ├── Data/
│   |   ├── Tesseract-OCR/
│   |   └── GUI DETS Correction System.py
│   └── Serial Number Maker/ # The code that automatic create serial numbers
│       ├── DETS Serial Number Maker.py
│       └── Students Data.xlsx
└── Templates/ # 10, 30, and 60 questions answer sheet templates
```

--- 

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests for improvements or bug fixes.

---

## License
This project is licensed under the [MIT License](https://github.com/barakota15/DETS-Correction-System/blob/main/LICENSE). See the LICENSE file for details.

---

## Acknowledgments

- Special thanks to the developers of [Tesseract OCR](https://github.com/tesseract-ocr/tesseract/blob/main/README.md) for their powerful OCR engine.
- Thanks to the contributors of the Python libraries used in this project.

---

## Author

SubNaut by Barakota15 — version 1.2
