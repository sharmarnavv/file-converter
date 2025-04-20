import os
import aspose.words as aw
import aspose.slides as slides
from tqdm import tqdm

def convert(path):
    try:
        if path.endswith(('.docx', '.doc')):
            doc = aw.Document(path)
            pdfPath = os.path.splitext(path)[0] + '.pdf'
            doc.save(pdfPath)
            print(f"Converted: {path}")
        elif path.endswith(('.pptx', '.ppt')):
            presentation = slides.Presentation(path)
            pdfPath = os.path.splitext(path)[0] + '.pdf'
            presentation.save(pdfPath, slides.export.SaveFormat.PDF)
            print(f"Converted: {path}")
    except Exception as e:
        print(f"Failed to convert {path}: {e}")

def collect_files(path, files):
    try:
        for item in os.listdir(path):
            item_path = os.path.join(path, item)
            if os.path.isfile(item_path):
                _, ext = os.path.splitext(item_path)
                if ext.lower() in ('.doc', '.docx', '.ppt', '.pptx'):
                    files.append(item_path)
            elif os.path.isdir(item_path):
                collect_files(item_path, files)
    except FileNotFoundError:
        print(f"Directory not found at '{path}'")

def driver(path, delete=False):
    files = []
    collect_files(path, files)

    if not files:
        print("No files found for conversion.")
        return

    for file in tqdm(files, desc="Converting Files"):
        convert(file)
        if delete:
            os.remove(file)

if __name__ == '__main__':
    d = input("Press 1 to delete files after conversion: ")
    delete = d.strip() == '1'
    folder_path = input("Enter directory path: ")
    driver(folder_path, delete)
