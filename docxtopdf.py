import os
import win32com.client
wdFormatPDF = 17
path = os.getcwd()
files = os.listdir(path)
files_docx = [f for f in files if f[-4:] == 'docx']
files_docx
#del(files_xlsx[-2])
n=0
for f in files_docx:
    inputFile =os.path.abspath(f)
    outputFile = os.path.abspath(f+".pdf")
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
    doc.Close()
    print(n)
    n+=1
