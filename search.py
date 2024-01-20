import os
import argparse
import logging
from docx import Document
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

# configure logger
log_format = '(%(asctime)s) [%(levelname)s] %(message)s'
log_file_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + '.log'
log_file_path = os.path.join(os.getcwd(), log_file_name)

logging.basicConfig(level=logging.DEBUG, format=log_format, handlers=[
    logging.FileHandler(log_file_path),
    logging.StreamHandler()
])

logger = logging.getLogger(__name__)

def check(fpath, target):
    try:
        doc = Document(fpath)
        for paragraph in doc.paragraphs:
            if target in paragraph.text:
                return True
        return False
    except Exception as e:
        logger.error(f"Error processing {fpath}: {e}")
        return False

def process_file(file):
    fname, target = file
    fpath = os.path.join(os.getcwd(), fname)
    if check(fpath, target):
        logger.info(f"'{target}' found in {fname}")
    else:
        logger.debug(f"'{target}' not found in {fname}")

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
    logger.debug(f"Execution Time: {execution_time} seconds")
