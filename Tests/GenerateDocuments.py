from HtmlToWord.parser.html import HTMLParser
from HtmlToWord.render.com import COMRenderer
import win32com.client
import os
import sys
import pprint

if not os.path.exists("saved_documents"):
    os.mkdir("saved_documents")

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = True
parser = HTMLParser()

try:
    paths = (sys.argv[1],)
except IndexError:
    paths = (path for path in os.listdir("html") if path.endswith(".html"))

for file_name in paths:
    print("Parsing: %s" % file_name)

    document = word.Documents.Add()

    renderer = COMRenderer(word, document, document.ActiveWindow.Selection)

    with open(os.path.join("html", file_name), "r") as fd:
        Html = fd.read()

    renderer.render(parser.parse(Html))

    path = os.path.abspath(os.path.join("saved_documents", file_name + ".docx"))

    document.SaveAs(path)
    #document.Close()