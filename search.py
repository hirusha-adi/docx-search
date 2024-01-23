import argparse
import os
from docx_search import docx_search

def main():
    parser = argparse.ArgumentParser(description='Search for a word in .docx files in a specified directory.')
    parser.add_argument(
        '--dir', dest='target_dir', type=str, default=os.getcwd(),
                        help='The target directory to search for .docx files. Defaults to the current working directory.'
                        )
    parser.add_argument('--word', dest='target_word', type=str,
                        help='The word to search for in the documents. If not provided, user will be prompted.')
    args = parser.parse_args()

    target_dir = args.target_dir
    target_word = args.target_word

    if not target_word:
        target_word = input("Enter the word to search for: ")

    docx_search(target_dir=target_dir, target_word=target_word)

if __name__ == "__main__":
    main()
