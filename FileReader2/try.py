import os
import subprocess as sp
import PyPDF2
from tqdm import trange, tqdm
import xlwings as xw
from itertools import chain
import re

FILE = "file/HaaS_v5.4.pdf"
page_lines = []
line_strip_list = []

CURR_DIR = os.getcwd()

XL_FILE = CURR_DIR+'/command_generator.xlsm'
sheet = xw.Book(XL_FILE).sheets("Sheet1")
user_input = input("Enter chapter section :")

chapters = {    
    "2.1":
        {   "start":"1 HaaSホスト名を設定します。",            
            "end": "1 ブリッジ （br-haas -mng）を作成します。"
        },
    "2.2":
        {   "start":"1 ブリッジ （br-haas -mng）を作成します。",            
            "end": "1 ブリッジ （br-haas -op）を作成します。"
        },
    "2.3":
        {   "start":"1 ブリッジ （br-haas -op）を作成します。",           
            "end": "Zabbix  Agentの設定手順を以下に示します。 "
        }                    
}

def parse(file_name: str) -> list:
    with open(file_name, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)

        for pn in trange(len(reader.pages), desc='Parsing data from PDF file',colour='white'):
            if pn>=14:
                text_page = reader.pages[pn].extract_text()                
                fil_page = text_page.split("\n")
                page_lines.append(fil_page)
                
        with open("output.txt", 'w', encoding='utf-8') as f:
            for data in page_lines:
                f.write('\n'.join(str(i) for i in data)+"\n")
                #print(str(data) +"\n")zzzz
   
def write_to_work_book():            
    with open("output.txt", "r", encoding='utf-8') as f:
        for key in chapters:                                      
            if key == user_input:   
                for line in f:             
                    start_val = chapters[key]['start']
                    end_val = chapters[key]['end']                                       
                    if line.rstrip() == start_val:                                                     
                        for line in f:
                            if line.rstrip() == end_val:
                                break
                            line_strip = line.rstrip()
                            line_strip_list.append(line_strip)
            if key != user_input:   
                print("Invalid chapter")
                break
                           
                          
    validation = re.compile(r'^#[^#|^A-Z][./|a-z]+') 
    
    if (line_strip_list):                    
        commands = [word for word in line_strip_list if re.match(validation, word)]     
                        
        row=1
        """Function to write to excel workbook sheet"""   
        sheet.range("A1:A200").clear_contents()      
        for idx, command in tqdm(enumerate(commands, start=row),desc='Writing to Excel sheet workbook',colour='green'):    
            sheet.range((idx, 1)).value=command
            row = idx+1
                     
def main():
    parse(FILE)
    write_to_work_book()


if __name__ == '__main__':
    main()
                


