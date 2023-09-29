import os
import logging
from docx import Document
from collections import Counter
from nltk import sent_tokenize

logging.basicConfig(level=logging.INFO)

def search_and_summarize(file_path, search_terms):
    try:
        doc = Document(file_path)
        term_counts = Counter()
        term_sentences = {term: [] for term in search_terms}

        for paragraph in doc.paragraphs:
            text = paragraph.text.lower()
            for term in search_terms:
                count = text.count(term.lower())
                term_counts[term] += count
                if count > 0:
                    sentences = sent_tokenize(text)
                    term_sentences[term].extend(sentences)

        return term_counts, term_sentences

    except FileNotFoundError as e:
        logging.error(f"The file does not exist: {e}")
        raise e
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        raise e

def is_docx_file(file_path):
    _, file_extension = os.path.splitext(file_path)
    return file_extension.lower() == '.docx'

def get_valid_input():
    while True:
        file_path = input("Enter the path to the .docx file: ")
        
        if not os.path.exists(file_path):
            print("The file does not exist. Please enter a valid file path.")
            continue
        
        if not is_docx_file(file_path):
            print("The file is not in .docx (Word Document) format. Please provide a valid .docx file.")
            continue
        
        search_terms = input("Enter one or more words or phrases to search for (comma-separated): ")
        if not search_terms:
            print("Please provide at least one search term.")
            continue
        
        return file_path, search_terms.split(",")

def main():
    try:
        file_path, search_terms = get_valid_input()
        term_counts, term_sentences = search_and_summarize(file_path, search_terms)
        
        print("===== Word/Phrase Counts =====")
        for term, count in term_counts.items():
            print(f"{term}: {count}")

        print("\n===== Here's the Summary Sentences =====")
        for term, sentences in term_sentences.items():
            print(f"{term} Sentences:")
            for sentence in sentences:
                print(sentence)
            print()

    except FileNotFoundError as e:
        logging.error(f"Error: {e}")
    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
