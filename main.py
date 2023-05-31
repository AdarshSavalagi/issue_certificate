import openpyxl
import ppt_generator


def get_rows(file_path, sheet_name):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]

    rows = []
    for row in sheet.iter_rows(values_only=True):
        rows.append(row)

    return rows


# def mail_function(details):
#     replacements = [('{name}', details[0]), ('{college}', details[1]), ]
#     ppt_file = r'/ppts/example.pptx'
#     modified_file_name = PPTtoPDF(os.getcwd() + '\\' + replace_ppt_text(ppt_file, replacements), details[0])
#     print(modified_file_name )
#     # send_email_with_ppt(details[2], modified_file_name)
#     # if os.path.exists(modified_file_name):
#     #     print('came')
#         # os.remove(file_path)


if __name__ == '__main__':
    file_path = r'E:\pythonProject\pythonProject\sample.xlsx'
    sheet_name = 'Sheet1'  # Name of the sheet

    rows = get_rows(file_path, sheet_name)
    for row in rows:
        replacements = [('{name}', row[0]), ('{college}', row[1])]
        print('Line read successfully ', replacements)
        ppt_generator.replace_ppt_text(replacements, row[2])
