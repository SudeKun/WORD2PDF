from docx2pdf import convert as convert2docx
from win32com import client as wc
import os

answer=""

while(True):
    # Load word document
    file_path=input("Please enter the location of doc or docx document, to quit press enter:")
    if file_path=="":
        break
    if ".doc" in file_path and not ".docx" in file_path:
        w=wc.Dispatch('Word.Application')
        doc=w.Documents.Open(os.path.abspath(file_path))
        doc.SaveAs(file_path.replace("doc","docx"),16)
        w.Quit
        file_path=file_path.replace("doc","docx")
    file_path=file_path.replace("\'","")
    file_path=file_path.replace("\"","")
    file=file_path.replace("\\"," ")
    res = file.split(" ")
    filename=""
    for x in res:
        if ".docx" in x:
            filename=x.replace(".docx","")
            path=file_path.replace(x,"")
            convert2docx(file_path, (path+filename+".pdf"))
            break
    print("It's done!!")
quit()