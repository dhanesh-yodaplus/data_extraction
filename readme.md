# 🧾 Unstructured Data Extraction and PDF Report Generation (Python Project)

## ✅ What This Covers

| Module                      | Description                                                                    |
| --------------------------- | ------------------------------------------------------------------------------ |
| 📥 Text-Based PDF Batch     | Extracts structured fields from multiple PDFs using `pdfplumber` + `rapidfuzz` |
| 📄 Single PDF to PDF Report | Parses a text-based PDF and generates a summary PDF using `reportlab`          |
| 🖼 OCR on Scanned PDFs       | Uses `pdf2image` + `pytesseract` to extract text from scanned/image-based PDFs |
| 📝 Word Document Extraction | Reads `.docx` files with `python-docx` and extracts field data to Excel        |

---

## 🔧 Technologies Used

- `pdfplumber` – Extract text from text-based PDFs
- `pytesseract` – OCR engine for scanned images
- `pdf2image` – Converts PDF pages to images
- `rapidfuzz` – Fuzzy matching for resilient text parsing
- `python-docx` – Parse Word `.docx` documents
- `reportlab` – Generate clean PDF reports programmatically
- `pandas` – Structure and export data to Excel
