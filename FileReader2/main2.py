import os, sys, glob, pathlib
from itertools import chain
import PyPDF2
from openpyxl import Workbook, load_workbook
from tqdm import trange
import json
import re


XL_FILE = 'file/Sample.xlsx'
book = load_workbook(XL_FILE)
sheet = book.active

CURR_DIR = os.getcwd()

FILE = 'file/HaaS_v5.4.pdf'
FILE_1 = 'csv/sample.csv'
FILE_2 = 'sample.json'

chapters = []

# TODO: Get By Chapters Not Pages
# TODO: Add pattern to get only commands that starts with a '#'
# TODO: Determine PDF manual file in machine


def is_word_file(file_path: str) -> bool:
    if file_path.endswith('.docx'):
        return True
    return False


def is_pdf_file(file_path: str) -> bool:
    if file_path.endswith('.pdf'):
        return True
    return False


def parse(file_name: str, json_name='commands.json') -> list:
    """Function to parse command text from pdf file"""
    pattern = re.compile(r'#*')

    try:
        if is_pdf_file(file_name):
            with open(file_name, 'rb') as pdf_file:
                reader = PyPDF2.PdfReader(pdf_file)

                for pn in trange(len(reader.pages), desc='Parsing data from PDF file',
                                 colour='white'):
                    text_page = reader.pages[pn].extract_text()
                    # Add all text lines containing '#' at the start, to the list
                    # Using filter method
                    fil_page = filter(lambda word: word.startswith("#"), text_page.split("\n"))
                    page = {pn: list(fil_page)}
                    chapters.append(page)
                # Convert to json format
                json_str = json.dumps(chapters, indent=2)
                # Write as json file
                with open(json_name, 'w') as json_f:
                    json_f.write(json_str)
        return chapters

    except FileNotFoundError as Fnf:
        print(Fnf)


def write_to_work_book(data: list[dict], file_path: str,) -> None:
    """Function to write to excel workbook sheet"""
    with open(file_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        pages_len = len(reader.pages)

    # Write to excel workbook file for each value in the dictionary
    for i in trange(pages_len, desc='Writing to Excel sheet workbook',
                    colour='green'):
        sheet.append([]); sheet.append([f'Page Number: {i}']); sheet.append([])
        for command in chain(*data[i].values()):
            sheet.append([command])
        book.save(XL_FILE)


def main():
    parsed_data = parse(FILE)
    write_to_work_book(parsed_data, FILE)


if __name__ == '__main__':
    main()
