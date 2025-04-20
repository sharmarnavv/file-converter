# File Converter

A simple script to help me deal with all those annoying file types I get when downloading stuff from my college website. Converts DOC/DOCX and PPT/PPTX files to PDF automatically.

## Why I Made This
Got tired of manually converting files one by one when downloading lecture materials and assignments. This script just goes through a folder and converts everything to PDF in one go.

## Requirements
- Python
- LibreOffice installed on your computer
- The packages listed in requirements.txt

## How to Use
1. Install Python if you don't have it
2. Install LibreOffice if you don't have it
3. Create a `.env` file in the project folder with:
   ```
   LIBREOFFICE_PATH=your_path_here
   ```
   Example for Windows: `LIBREOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe`
   Example for Mac: `LIBREOFFICE_PATH=/Applications/LibreOffice.app/Contents/MacOS/soffice`
4. Run `pip install -r requirements.txt`
5. Run `python main.py`
6. Enter the folder path where your files are
7. Choose if you want to keep or delete the original files
8. That's it! All your files will be converted to PDF

## What It Converts
- Word docs (.doc, .docx)
- PowerPoint presentations (.ppt, .pptx)

All files get converted to PDF format because PDFs are life.
