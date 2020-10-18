import docx
from win32com.client import Dispatch
import os
import traceback
import textract
import re
import string
import pandas as pd
import matplotlib.pyplot as plt



def excel_pre():
    '''Excel Path Setting'''
    global xl
    xl = Dispatch("Excel.Application")
    xl.Visible = False #True = Display， False = Hide
    xl.DisplayAlerts = 0

def doc2Docx(fileName):
    '''Doc Convert to Docx'''
    word = Dispatch("Word.Application")
    doc = word.Documents.Open(fileName)
    doc.SaveAs(fileName + "x", 12, False, "", True, "", False, False, False, False)
    os.remove(fileName)
    doc.Close()
    word.Quit()

def get_text(docname):
    doc = docx.Document(cv_file_path + docname)
    fulltext = []
    for i, paragraph in enumerate(doc.paragraphs):
        this_text = paragraph.text
        this_text = this_text.lower()
        this_text = re.sub(r'\d+','',text)
        this_text = text.translate(str.maketrans('','',string.punctuation))

        if '[ Candidate Name ]' in this_text:
            this_text = this_text.split(' : ')[1]
            cv_dict['Candidate Name'] = this_text
        elif '[ Position Applied ]' in this_text:
            this_text = this_text.split(' : ')[1]
            cv_dict['Position Applied'] = this_text
        elif '[ Current Package ]' in this_text:
            this_text = this_text.split(' : ')[1]
            cv_dict['Current Package'] = this_text
        elif '[ Expected Package ]' in this_text:
            this_text = this_text.split(' : ')[1]
            cv_dict['Expected Package'] = this_text
        elif '[ Reason(s) for Leaving / Applying ]' in this_text:
            this_text = this_text.split(' : ')[1]
            cv_dict['Reasons for Leaving'] = this_text
        elif '[ Reason (s) for Leaving / Applying ]' in this_text:
            this_text = this_text.split(' : ')[1]
            cv_dict['Reasons for Leaving'] = this_text
        elif '[ Availability ]' in this_text:
            this_text = this_text.split(' : ')[1]
            cv_dict['Availability'] = this_text
    return score

def make_newxls(cv_dict, newrow):
    '''Export CV Info to Excel'''
    ws.Cells(newrow, 1).Value = cv_dict['CV Date']
    ws.Cells(newrow, 2).Value = cv_dict['Candidate Name']
    ws.Cells(newrow, 3).Value = cv_dict['Position Applied']
    ws.Cells(newrow, 4).Value = cv_dict['Current Package']
    ws.Cells(newrow, 5).Value = cv_dict['Expected Package']
    ws.Cells(newrow, 6).Value = cv_dict['Reasons for Leaving']
    ws.Cells(newrow, 7).Value = cv_dict['Availability']

if __name__ == "__main__":
    cv_file_path = os.path.abspath('.') + '\\' + 'CV Database' + '\\'
    for f in os.listdir('CV Database'):
        if f.endswith('.doc'):
            doc2Docx(cv_file_path + f)
    this_path = os.path.abspath('.') + '\\'
    excel_pre()
    wb = xl.Workbooks.Open(this_path + 'Contact Database.xls')
    ws = wb.Sheets('Contacts')
    n = 2
    try:
        for f in os.listdir('CV Database'):
            if f.endswith('.docx'):
                print(f)
                cv_dict = get_cv_info(f)
                make_newxls(cv_dict, n) #把记录放在《Contact Database.xls》里
                n +=1
    except:
        traceback.print_exc()
    finally:
        wb.Save()
        wb.Close()
