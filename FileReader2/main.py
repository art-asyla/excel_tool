import os, sys, glob, pathlib
from itertools import chain
import PyPDF2
from openpyxl import Workbook, load_workbook
from tqdm import trange
import json
import re
import xlwings as xw

CURR_DIR = os.getcwd()

# os.chdir(os.path.dirname(sys.executable))

XL_FILE = CURR_DIR+'/command_generator.xlsm'
sheet = xw.Book(XL_FILE).sheets("Sheet1")
# book = load_workbook(XL_FILE)
# sheet = book.active


FILE = CURR_DIR+'/file/HaaS_v5.4.pdf'
FILE_1 = 'csv/sample.csv'
FILE_2 = 'sample.json'

chapters = []

# TODO: Get By Chapters Not Pages
# TODO: Add pattern to get only commands that starts with a '#'
# TODO: Determine PDF manual file in machine


# def is_word_file(file_path: str) -> bool:
#     if file_path.endswith('.docx'):
#         return True
#     return False


def is_pdf_file(file_path: str) -> bool:
    if file_path.endswith('.pdf'):
        return True
    return False


def parse(file_name: str, json_name='commands.json') -> list:
    """Function to parse command text from pdf file"""
    #pattern = re.compile(r'#?')    

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
                    page = {f'Page {pn}': list(fil_page)}
                    chapters.append(page)
                # Convert to json format
                json_str = json.dumps(chapters, indent=2)
                # Write as json file
                with open(json_name, 'w') as json_f:
                    json_f.write(json_str)
        return chapters

    except FileNotFoundError as Fnf:
        print(Fnf)


def write_to_work_book(data) -> None:
    row=1
    """Function to write to excel workbook sheet"""
    # Write to excel workbook file for each value in the dictionary
    for i in trange(len(data), desc='Writing to Excel sheet workbook',
                    colour='green'):       
        #sheet.append([]); sheet.append([f'Page Number: {i+1}']); sheet.append([])
        sheet.range((row+1, 1)).value = ""
        sheet.range((row+2, 1)).value = f'Page Number: {i+1}'
        sheet.range((row+3, 1)).value = ""
        
        for idx, command in enumerate(chain(*data[i].values()), start=row+4):
            sheet.range((idx, 1)).value=command
            row = idx+1
            #sheet.append([command])
        #book.save(XL_FILE)


def main():
    parsed_data = parse(FILE)
    #print(parsed_data)
    write_to_work_book(parsed_data)


if __name__ == '__main__':
    main()
