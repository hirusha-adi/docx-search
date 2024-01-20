import os
from docx import Document

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

def main():
    target = "hirusha"
    cwd = os.getcwd()
    for fname in os.listdir(cwd):
        if fname.endswith(".docx"):
            fpath = os.path.join(cwd, fname)
            if search(fpath=fpath, target=target):
                print("Found in ", fname)
            else:
                print("Not found in ", fname)

if __name__ == "__main__":
    main()
