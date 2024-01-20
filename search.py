import os
from docx import Document
from concurrent.futures import ThreadPoolExecutor

def search(fpath, target):
    try:
        doc = Document(fpath)
        for paragraph in doc.paragraphs:
            if target in paragraph.text:
                return True
        return False
    except Exception as e:
        print(e)
        return False

def process_file(file):
    fname, target = file
    fpath = os.path.join(os.getcwd(), fname)
    if search(fpath=fpath, target=target):
        print("Found in ", fname)
    else:
        print("Not found in ", fname)
        
def main():
    target = "hirusha"
    all_files = [(fname, target) for fname in os.listdir() if fname.endswith(".docx")]
    
    with ThreadPoolExecutor() as  executor:
        executor.map(process_file, all_files)

if __name__ == "__main__":
    main()
