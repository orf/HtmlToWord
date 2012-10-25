import HtmlToWord
import win32com.client
import os
import sys

if not os.path.exists("saved_documents"):
    os.mkdir("saved_documents")

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = True
parser = HtmlToWord.Parser()

try:
    paths = (sys.argv[1],)
except IndexError:
    paths = (path for path in os.listdir("html") if path.endswith(".html"))

for file_name in paths:
    document = word.Documents.Add()

    with open(os.path.join("html",file_name),"r") as fd:
        Html = ""
        for line in fd:
            line = line.replace("\n","")
            line = line.rstrip().lstrip()
            Html+=line
    print parser.Parse(Html)
    parser.ParseAndRender(Html, word, document.ActiveWindow.Selection)
    path = os.path.abspath(os.path.join("saved_documents",file_name+".docx"))
    print path
    document.SaveAs(path)
    document.Close()