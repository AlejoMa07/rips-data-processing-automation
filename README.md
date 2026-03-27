# 📊 RIPS Data Processing Automation

Automates the extraction, transformation, and loading of healthcare data (RIPS) into structured Excel templates using OCR and data validation techniques.

---

## 🚨 Problem

Processing RIPS-related documents required manual data extraction, validation, and entry into Excel templates, resulting in high time consumption when handling large volumes of records.

---

## ⚙️ Solution

A Python-based automation pipeline was developed to extract information from PDF documents using OCR, process and validate the data, and automatically populate structured Excel templates ready for upload.

---

## 🔍 Features

* OCR-based data extraction from PDF files
* Automatic identification of key fields (ID, name, dates, city, etc.)
* Data normalization and validation
* Automatic Excel template population
* Handling of multiple documents in batch
* Integration with reference tables for data matching
* Automatic file and folder organization

---

## ▶️ Workflow

1. Read PDF files from input folder
2. Extract relevant data using OCR
3. Process and normalize extracted information
4. Validate fields against reference data
5. Populate structured Excel templates
6. Generate outputs ready for system upload

---

## 📁 Project Structure

* `/input` → PDF files to process
* `/output` → generated Excel files
* `rips_processor.py` → main automation script
* `README.md` → project documentation

---

## 📈 Results

* Significant reduction in manual data entry time
* Ability to process large volumes of documents efficiently
* Standardization of data for system compatibility
* Improved operational efficiency in RIPS processing

---

## 🛠️ Technologies

* Python
* Tesseract OCR
* pdf2image
* OpenPyXL
* xlwings
* Fuzzy matching (thefuzz)

---

## 📌 Author

Luis Alejandro Machado
