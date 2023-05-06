import random
import string
import os
from docx import Document
import openpyxl
import tkinter as tk
from tkinter import filedialog

def generate_random_string(length):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

def anonymize_docx(input_file, output_file, sensitive_terms):
    document = Document(input_file)
    anonymized_terms = {}

    for term in sensitive_terms:
        placeholder = generate_random_string(len(term))
        anonymized_terms[term] = placeholder

        for paragraph in document.paragraphs:
            if term in paragraph.text:
                paragraph.text = paragraph.text.replace(term, placeholder)

    document.save(output_file)
    return anonymized_terms

def anonymize_xlsx(input_file, output_file, sensitive_terms):
    workbook = openpyxl.load_workbook(input_file)
    anonymized_terms = {}

    for term in sensitive_terms:
        placeholder = generate_random_string(len(term))
        anonymized_terms[term] = placeholder

        for sheet in workbook:
            for row in sheet.iter_rows():
                for cell in row:
                    if term in str(cell.value):
                        cell.value = str(cell.value).replace(term, placeholder)

    workbook.save(output_file)
    return anonymized_terms

def save_mapping_to_file(anonymized_terms, mapping_file):
    with open(mapping_file, 'w', encoding='utf-8') as f:
        for original, anonymized in anonymized_terms.items():
            f.write(f"{original} -> {anonymized}\n")

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path

sensitive_terms = ['Nick']

input_file = select_file()
output_file = os.path.splitext(input_file)[0] + '_anonymized' + os.path.splitext(input_file)[1]
mapping_file = 'anonymized_terms_mapping.txt'

file_extension = os.path.splitext(input_file)[1]

if file_extension == '.docx':
    anonymized_terms = anonymize_docx(input_file, output_file, sensitive_terms)
elif file_extension == '.xlsx':
    anonymized_terms = anonymize_xlsx(input_file, output_file, sensitive_terms)
else:
    raise ValueError("Unsupported file format. Only .docx and .xlsx files are supported.")

save_mapping_to_file(anonymized_terms, mapping_file)
print("Anonymization process completed. Check the output files.")
