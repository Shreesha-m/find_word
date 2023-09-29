import os
import logging
import win32com.client
from collections import Counter
from nltk import sent_tokenize

# Set up logging to display INFO level messages
logging.basicConfig(level=logging.INFO)

def search_and_summarize(file_path, search_terms):
    try:
        # Create a COM object for Microsoft Word
        word = win32com.client.Dispatch("Word.Application")
        # Open the specified .doc or .docx file
        doc = word.Documents.Open(file_path)
        # Initialize counters and sentence storage
        term_counts = Counter()
        term_sentences = {term: [] for term in search_terms}

        # Iterate through paragraphs in the document
        for paragraph in doc.Paragraphs:
            text = paragraph.Range.Text.lower()
            for term in search_terms:
                count = text.count(term.lower())
                term_counts[term] += count
                if count > 0:
                    sentences = sent_tokenize(text)
                    term_sentences[term].extend(sentences)

        # Close the document and quit Microsoft Word
        doc.Close()
        word.Quit()

        return term_counts, term_sentences

    except FileNotFoundError as e:
        logging.error(f"The file does not exist: {e}")
        raise e
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        raise e

def is_doc_file(file_path):
    _, file_extension = os.path.splitext(file_path)
    return file_extension.lower() == '.doc'

def get_valid_input():
    while True:
        file_format = input("Choose the document format (1 for .doc, 2 for .docx, or 'q' to quit): ")
        
        if file_format.lower() == 'q':
            exit(0)  # Exit the program if the user enters 'q'
        
        if file_format not in ('1', '2'):
            print("Invalid choice. Please enter '1' for .doc or '2' for .docx.")
            continue
        
        file_extension = '.doc' if file_format == '1' else '.docx'
        
        file_path = input(f"Enter the path to the document file ({file_extension}) or 'q' to quit: ")
        
        if file_path.lower() == 'q':
            exit(0)  # Exit the program if the user enters 'q'
        
        if not os.path.exists(file_path):
            print("The file does not exist. Please enter a valid file path.")
            continue
        
        search_terms = input("Enter one or more words or phrases to search for (comma-separated): ")
        if not search_terms:
            print("Please provide at least one search term.")
            continue
        
        return file_path, search_terms.split(",")

def main():
    while True:
        try:
            file_path, search_terms = get_valid_input()
            term_counts, term_sentences = search_and_summarize(file_path, search_terms)
            
            # Print word/phrase counts
            print("===== Word/Phrase Counts =====")
            for term, count in term_counts.items():
                print(f"{term}: {count}")

            # Print summary sentences
            print("\n===== Here's the Summary Sentences =====")
            for term, sentences in term_sentences.items():
                if term_counts[term] > 0:
                    print(f"\n{term} Sentences:")
                    # Limit the number of displayed sentences (adjust this as needed)
                    max_sentences = 5
                    for i, sentence in enumerate(sentences):
                        if i >= max_sentences:
                            print(f"... and {len(sentences) - max_sentences} more sentences.")
                            break
                        print(f"  - {sentence}")

            repeat = input("Do you want to perform another search? (yes/no): ").lower()
            if repeat != 'yes':
                break  # Exit the loop if the user doesn't want to repeat

        except FileNotFoundError as e:
            logging.error(f"Error: {e}")
        except Exception as e:
            logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
