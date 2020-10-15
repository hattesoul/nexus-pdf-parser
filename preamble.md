# Prerequisite

1. The documents were scanned Nexus PDF reports
2. The PDF files were OCR-processed with OCRmyPDF
    1. install OCRmyPDF: `sudo apt install ocrmypdf`
    2. install German for Tesseract `sudo apt install tesseract-ocr-deu`
    3. convert all PDF in the current folder: `for f in ./*.pdf; do ocrmypdf -f -l deu "$f" "$(basename "$f" ".pdf")_ocr.pdf"; done`
