import os
import subprocess
from dotenv import load_dotenv
from tqdm import tqdm

load_dotenv()

LIBREOFFICE_PATH = os.getenv("LIBREOFFICE_PATH")

if not LIBREOFFICE_PATH or not os.path.isfile(LIBREOFFICE_PATH):
    print("❌ Error: LIBREOFFICE_PATH is not set correctly in your .env file.")
    print("Expected an absolute path to soffice.exe, like: D:\\files\\LibreOffice\\program\\soffice.exe")
    exit(1)

def convert(input_file):
    try:
        subprocess.run([
            LIBREOFFICE_PATH,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", os.path.dirname(input_file),
            input_file
        ], check=True)
        print(f"\nConverted: {input_file}")
    except subprocess.CalledProcessError as e:
        print(f"\nFailed to convert {input_file}: {e}")

def collect_files(path, extensions):
    collected = []
    for root, _, files in os.walk(path):
        for file in files:
            if file.lower().endswith(extensions):
                collected.append(os.path.join(root, file))
    return collected

def driver(path, delete=False):
    files = collect_files(path, ('.docx', '.doc', '.pptx', '.ppt'))
    if not files:
        print("No files to convert.")
        return

    for file in tqdm(files, desc="Converting"):
        convert(file)
        if delete:
            os.remove(file)

if __name__ == '__main__':
    d = input("Press 1 to delete files after conversion: ")
    delete = d.strip() == '1'
    folder_path = input("Enter directory path: ")
    driver(folder_path, delete)
