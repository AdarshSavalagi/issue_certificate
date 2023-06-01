import os
import shutil
import openpyxl
import WORD_TO_PDF
import sendmail


def get_rows(file_path, sheet_name):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    rows = []
    for row in sheet.iter_rows(values_only=True):
        rows.append(row)
    return rows


if __name__ == '__main__':
    file_path = r'E:\pythonProject\issue_certificate\sample.xlsx'
    sheet_name = 'Sheet1'

    rows = get_rows(file_path, sheet_name)
    for row in rows:
        replacements = [row[0], row[1]]
        print('Line read successfully ', replacements)
        pdf_path = WORD_TO_PDF.replace_text_in_word(replacements)
        sendmail.send_mail(row[2], pdf_path)
    os.remove(r'E:\pythonProject\issue_certificate\ppts\cert_new.docx')
    shutil.rmtree(r'E:\pythonProject\issue_certificate\pdfs')
    os.mkdir(r'E:\pythonProject\issue_certificate\pdfs')
