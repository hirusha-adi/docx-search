# docx-search

**Description:**
The `docx-search` Python script is a tool designed to search for a specified word within Microsoft Word (.docx) documents in a given directory. It utilizes the `python-docx` library for handling Word documents and implements multi-threading to improve search efficiency.

**Features:**

- **Word Search:** The script searches for a specified target word within the paragraphs of each Word document in the provided directory.
- **Logging:** Detailed logging is implemented, capturing information about the search process, including the presence or absence of the target word in each document.
- **Multi-threading:** The script utilizes the `concurrent.futures.ThreadPoolExecutor` to concurrently process multiple Word documents, improving overall search performance.

**Getting Started:**

1. **Requirements:**

   - Python 3.x
   - Install required Python packages using `pip install python-docx`

2. **Usage:**

   - Run the script from the command line:
     ```
     python search.py
     ```
   - The script will search for the specified target word in all `.docx` files within the current working directory by default.

   - Optionally, you can specify the target directory and word using command-line arguments:
     ```
     python search.py --dir /path/to/documents --word example
     ```

3. **Logging:**

   - Logs are saved in the 'logs' directory with filenames in the format 'YYYY-MM-DD_HH-MM-SS.log.'

4. **Output:**

   - The script outputs information about the presence or absence of the target word in each processed document.

5. **Execution Time:**
   - The script logs the execution time, providing insights into the performance of the search operation.

**Using as a Module:**

1. **Import the Module:**

   - Import the `docx_search` module into your Python script:
     ```python
     from docx_search import docx_search
     ```

2. **Perform Word Search:**
   - Call the `docx_search` function with the desired target directory and word:
     ```python
     docx_search(target_dir="/path/to/documents", target_word="example")
     ```

**Notes:**

- Ensure the `python-docx` library is installed before running the script.

**Contributing:**

- Contributions are welcome! Feel free to fork the repository, make improvements, and create a pull request.

**License:**

- This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

**Acknowledgments:**

- This readme.md and the docstrings were generated with ChatGPT, a language model developed by OpenAI.
