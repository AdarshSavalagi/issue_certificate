import win32com.client as win32


def replace_text_in_word(new_texts):
    # Create an instance of the Word application
    word_app = win32.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False  # Set to True if you want to see Word in action
    file_path = r'E:\pythonProject\issue_certificate\ppts\cert.docx'
    old_texts = ['name ', 'college']

    # Open the Word document
    doc = word_app.Documents.Open(file_path)
    doc.SaveAs(r'E:\pythonProject\issue_certificate\ppts\cert_new.docx')
    doc.Close()
    file_path = r'E:\pythonProject\issue_certificate\ppts\cert_new.docx'
    doc = word_app.Documents.Open(file_path)

    # Replace text in main document
    for old_text, new_text in zip(old_texts, new_texts):
        doc.Content.Find.Text = old_text
        doc.Content.Find.Replacement.Text = new_text
        doc.Content.Find.Execute(Replace=2)

    # Replace text in text boxes
    for shape in doc.Shapes:
        if shape.Type == 17:  # Check if shape is a text box
            if shape.TextFrame.HasText:
                for old_text, new_text in zip(old_texts, new_texts):
                    shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.replace(old_text, new_text)

    # Save the modified document as PDF
    pdf_path = rf'E:\pythonProject\issue_certificate\pdfs\{new_texts[0]}.pdf'
    doc.SaveAs(pdf_path, FileFormat=17)

    # Close the document
    doc.Close()

    # Quit Word application
    word_app.Quit()

    return pdf_path

# Example usage
# file_path = r'E:\pythonProject\issue_certificate\ppts\cert.docx'
# old_texts = ['name ', 'college']
# new_texts = ['Hi there', 'Python']
#
# pdf_path = replace_text_in_word(file_path, old_texts, new_texts)
# print("PDF saved at:", pdf_path)
