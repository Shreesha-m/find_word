# Document Search and Summarization Script

This Python script allows you to search for specified words or phrases within Microsoft Word documents (`.doc` or `.docx`) and provides a summary of the counts of these words/phrases and their occurrences in sentences.

## Prerequisites

Before using the script, ensure you have the following installed:

- Python 3.x
- Required Python packages (install using `pip`):
  - `win32com.client`
  - `nltk`

You can install the required packages using the following command:

```bash
pip install pypiwin32 nltk

......
Usage
1. Clone or download this repository to your local machine.

2. Navigate to the directory containing the script.

3. Run the script using the following command:
   python find_word.py
4. Follow the on-screen prompts to perform a search in your Microsoft Word document

Script Functionality
1. The script allows you to choose the document format (.doc or .docx) and specify the search terms.

2. It displays the counts of each search term and provides summary sentences where each term appears in the document.

3. You can perform multiple searches in succession or exit the script when you're done.

Example
Here's an example of using the script:

1. Choose the document format (1 for .doc, 2 for .docx).
2. Enter the path to your document file.
3. Enter one or more search terms (comma-separated).
4. View the word/phrase counts and summary sentences.
5. Choose whether to perform another search.

License
This project is licensed under the MIT License - see the LICENSE.md file for details.

Acknowledgments
Python
pypiwin32
NLTK
