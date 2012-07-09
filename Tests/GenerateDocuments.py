import HtmlToWord
from HtmlToWord.elements import *
import win32com.client
import os
if not os.path.exists("saved_documents"):
    os.mkdir("saved_documents")

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = True
parser = HtmlToWord.Parser()

for file_name in (path for path in os.listdir("html") if path.endswith(".html")):
    document = word.Documents.Add()

    with open(os.path.join("html",file_name),"r") as fd:
        Html = ""
        for line in fd:
            line = line.replace("\n","")
            line = line.rstrip().lstrip()
            Html+=line

    parser.ParseAndRender(Html, word, document.ActiveWindow.Selection)
    path = os.path.abspath(os.path.join("saved_documents",file_name+".docx"))
    print path
    document.SaveAs(path)
    document.Close()