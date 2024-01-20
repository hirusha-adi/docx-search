import os
import argparse
from docx import Document
from concurrent.futures import ThreadPoolExecutor

def check(fpath, target):
    try:
        doc = Document(fpath)
        for paragraph in doc.paragraphs:
            if target in paragraph.text:
                return True
        return False
    except Exception as e:
        print(f"Error processing {fpath}: {e}")
        return False

def process_file(file):
    fname, target = file
    fpath = os.path.join(os.getcwd(), fname)
    if check(fpath, target):
        print(f"'{target}' found in {fname}")
    else:
        print(f"'{target}' not found in {fname}")

def main():
    parser = argparse.ArgumentParser(description='Search for a word in .docx files in the current directory.')
    parser.add_argument('word', type=str, help='The word to search for')
    args = parser.parse_args()

    target = args.word
    fall = [(fname, target) for fname in os.listdir() if fname.endswith(".docx")]

    with ThreadPoolExecutor() as executor:
        executor.map(process_file, fall)

if __name__ == "__main__":
    import time
    start_time = time.time()
    main()  
    end_time = time.time()
    execution_time = end_time - start_time
    print(f"Execution Time: {execution_time} seconds")
    