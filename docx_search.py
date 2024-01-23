import os
import argparse
import logging
from docx import Document
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

# Configure logger
log_format = '(%(asctime)s) [%(levelname)s] %(message)s'
log_file_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + '.log'
log_file_path = os.path.join(os.getcwd(),"logs", log_file_name)

logging.basicConfig(level=logging.DEBUG, format=log_format, handlers=[
    logging.FileHandler(log_file_path, encoding='utf-8'),
    logging.StreamHandler()
])

logger = logging.getLogger(__name__)

def __check(fpath, target):
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

def __process_file(file):
    """
    Process a Word document file, checking if a target word is present.

    @param file: A tuple containing the file name and the target word.
    @type file: tuple
    @return: None
    @rtype: None
    """
    fname, target = file
    fpath = os.path.join(os.getcwd(), fname)
    if __check(fpath, target):
        logger.info("'%s' found in %s" % (target, fname))
    else:
        logger.debug("'%s' not found in %s" % (target, fname))

def __main(target_dir=None, target_word=None):
    """
    Main function to search for a word in .docx files in the current directory.

    @return: None
    @rtype: None
    """
    parser = argparse.ArgumentParser(description='Search for a word in .docx files in the current directory.')
    parser.add_argument('word', nargs='?', type=str, help='The word to search for (optional if not provided, will prompt user)')
    args = parser.parse_args()

    if target_word is not None and target_word != '':
        target = target_word
    elif args.word is not None and args.word != '':
        target = args.word
    else:
        raise ValueError("Both target_word and args.word are None or empty strings. Please pass in correct values.")

    file_list = [(fname, target) for fname in os.listdir(target_dir) if fname.endswith(".docx")]

    with ThreadPoolExecutor() as executor:
        executor.map(__process_file, file_list) # හිරුෂඅදිකාරී
        
def docx_search(target_dir=None, target_word=None):
    """
    Search for a target word in .docx files within a specified directory.

    This function calls the __main function to perform the actual search operation.
    
    @param target_dir: The target directory to search for .docx files. If None, the current working directory is used.
    @type target_dir: str or None
    @param target_word: The word to search for in the documents.
    @type target_word: str or None
    @return: None
    @rtype: None
    """
    import time
    start_time = time.time()
    __main(target_dir=target_dir, target_word=target_word)
    end_time = time.time()
    execution_time = end_time - start_time
    logger.debug("Execution Time: %s seconds" % execution_time)

if __name__ == "__main__":
    docx_search()