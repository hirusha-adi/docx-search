import os
import argparse
import logging
from docx import Document
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

class WordDocumentSearch:
    def __init__(self, target_word):
        self.target_word = target_word
        self.setup_logger()

    def setup_logger(self):
        log_format = '(%(asctime)s) [%(levelname)s] %(message)s'
        log_file_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + '.log'
        log_file_path = os.path.join(os.getcwd(), log_file_name)

        logging.basicConfig(level=logging.DEBUG, format=log_format, handlers=[
            logging.FileHandler(log_file_path),
            logging.StreamHandler()
        ])

        self.logger = logging.getLogger(__name__)

    def check_word_in_document(self, fpath):
        try:
            doc = Document(fpath)
            for paragraph in doc.paragraphs:
                if self.target_word in paragraph.text:
                    return True
            return False
        except Exception as e:
            self.logger.error("Error processing %s: %s" % (fpath, e))
            return False

    def process_file(self, file):
        fname, target = file
        fpath = os.path.join(os.getcwd(), fname)
        if self.check_word_in_document(fpath):
            self.logger.info("'%s' found in %s" % (self.target_word, fname))
        else:
            self.logger.debug("'%s' not found in %s" % (self.target_word, fname))

    def search_word_in_documents(self):
        file_list = [(fname, self.target_word) for fname in os.listdir() if fname.endswith(".docx")]

        with ThreadPoolExecutor() as executor:
            executor.map(self.process_file, file_list)

def main():
    parser = argparse.ArgumentParser(description='Search for a word in .docx files in the current directory.')
    parser.add_argument('word', type=str, help='The word to search for')
    args = parser.parse_args()

    word_search = WordDocumentSearch(args.word)
    word_search.search_word_in_documents()

if __name__ == "__main__":
    import time
    start_time = time.time()
    main()
    end_time = time.time()
    execution_time = end_time - start_time
    logging.getLogger(__name__).debug("Execution Time: %s seconds" % execution_time)
