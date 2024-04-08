import os
import re
import comtypes.client
from openpyxl import Workbook
from openpyxl.styles import Font
import docx
from openpyxl.utils import escape

from pdfminer.high_level import extract_text

from docx import Document
import comtypes

# Function to extract text from PDF using pdfminer


def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as f:
        raw_text = extract_text(f)

    cnt = 0
    final_text = ""
    for i in range(0, len(raw_text)):
        if (raw_text[i] == " " or raw_text[i] == "\n"):
            cnt += 1
        else:
            cnt = 0
        if (cnt <= 5):
            final_text += raw_text[i]
    return final_text.strip()

# Function to extract text from DOCX


def convert_doc_to_pdf(doc_name):

    word_path = doc_name
    pdf_path = doc_name.split(".")[0]+".pdf"

    folder_path = "Sample2"

    print(word_path, pdf_path)
    # word = comtypes.client.CreateObject('Word.Application')
    word = comtypes.client.CreateObject("Word.Application")
    docx_path = os.path.abspath(os.path.join(folder_path, word_path))
    pdf_path = os.path.abspath(os.path.join(folder_path, pdf_path))

    print(docx_path, pdf_path)

    pdf_format = 17
    word.Visible = False
    in_file = word.Documents.Open(docx_path)
    in_file.SaveAs(pdf_path, FileFormat=pdf_format)
    in_file.Close()
    word.Quit()


# Function to extract phone numbers


def extract_phone_numbers(text):
    phone_regex = re.compile(
        r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})')
    phone_numbers = re.findall(phone_regex, text)
    return phone_numbers

# Function to extract email addresses


def extract_emails(text):
    email_regex = re.compile(
        r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b')
    emails = re.findall(email_regex, text)
    return emails


# Helper function to filter out non-allowed characters
def filter_allowed_characters(text, allowed_characters):
    filtered_text = ""
    for char in text:
        if char in allowed_characters:
            filtered_text += char
        else:
            filtered_text += " "  # Replace non-allowed characters with space
    return filtered_text

# Function to write extracted data to Excel


def write_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active
    # Adding column header
    ws.append(["Name", "Phone Numbers", "Emails", "Text Extracted"])
    # Define a set of allowed characters
    allowed_characters = set(
        "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~ ")

    for entry in data:
        name, phone_numbers, emails, text_extracted = entry  # Unpacking the entry
        # Filter out non-allowed characters
        new_text = filter_allowed_characters(
            text_extracted, allowed_characters)

        # Convert phone numbers list to string
        phone_numbers_str = ", ".join(phone_numbers)

        # Appending to Excel sheet
        ws.append([name, phone_numbers_str, ", ".join(emails), new_text])

    wb.save(output_file)


# Main function


def main():
    data = []
    folder_path = "Sample2"
    output_file = "extracted_data.xlsx"

    for filename in os.listdir(folder_path):

        filename_in_pdf = filename.split(".")[0]+'.pdf'
        if (filename_in_pdf in os.listdir(folder_path)):
            continue
        if filename.endswith(".doc") or filename.endswith(".docx"):
            convert_doc_to_pdf(filename)

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith(".pdf"):
            text = extract_text_from_pdf(file_path)

            phone_numbers = extract_phone_numbers(text)
            emails = extract_emails(text)

            # Extract name from filename
            name = os.path.splitext(filename)[0]

            print(name, phone_numbers, emails)

            # Appending text extracted to data
            data.append([name, phone_numbers, emails, text])

    write_to_excel(data, output_file)
    print("Extraction completed. Data saved to", output_file)


if __name__ == "__main__":
    main()
