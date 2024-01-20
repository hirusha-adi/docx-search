import os
import argparse
import logging
from docx import Document
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

# Configure logger
log_format = '(%(asctime)s) [%(levelname)s] %(message)s'
log_file_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + '.log'
log_file_path = os.path.join(os.getcwd(), log_file_name)

logging.basicConfig(level=logging.DEBUG, format=log_format, handlers=[
    logging.FileHandler(log_file_path),
    logging.StreamHandler()
])

logger = logging.getLogger(__name__)

def check(fpath, target):
    """
    Check if the target word is present in the paragraphs of a given Word document.

    @param fpath: The file path of the Word document.
    @type fpath: str
    @param target: The word to search for in the document.
    @type target: str
    @return: True if the word is found, False otherwise.
    @rtype: bool
    """
    try:
        doc = Document(fpath)
        for paragraph in doc.paragraphs:
            if target in paragraph.text:
                return True
        return False
    except Exception as e:
        logger.error("Error processing %s: %s" % (fpath, e))
        return False

def process_file(file):
    """
    Process a Word document file, checking if a target word is present.

    @param file: A tuple containing the file name and the target word.
    @type file: tuple
    @return: None
    @rtype: None
    """
    fname, target = file
    fpath = os.path.join(os.getcwd(), fname)
    if check(fpath, target):
        logger.info("'%s' found in %s" % (target, fname))
    else:
        logger.debug("'%s' not found in %s" % (target, fname))

def main():
    """
    Main function to search for a word in .docx files in the current directory.

    @return: None
    @rtype: None
    """
    parser = argparse.ArgumentParser(description='Search for a word in .docx files in the current directory.')
    parser.add_argument('word', type=str, help='The word to search for')
    args = parser.parse_args()

    target = args.word
    file_list = [(fname, target) for fname in os.listdir() if fname.endswith(".docx")]

    with ThreadPoolExecutor() as executor:
        executor.map(process_file, file_list)

if __name__ == "__main__":
    import time
    start_time = time.time()
    main()
    end_time = time.time()
    execution_time = end_time - start_time
    logger.debug("Execution Time: %s seconds" % execution_time)
