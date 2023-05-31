import os
import comtypes.client
import sendmail


def ppt_to_pdf(inp, out, email):
    input_path = os.getcwd() + fr'\ppts\{inp}'
    output_path = os.getcwd() + fr'\pdfs\{out}'
    # try:
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    slides = powerpoint.Presentations.Open(input_path)
    slides.SaveAs(output_path, 32)  # 32 represents the PDF file format
    slides.Close()
    powerpoint.Quit()
    print("pdf generate aytu")
    sendmail.send_mail(email, fr'pdfs\{out}')
    # except Exception as e:
    #     print(e)
