from wordconverter.parser.html import HTMLParser
from wordconverter.parser.markdown import MarkdownParser
from wordconverter.render.com import COMRenderer
import win32com.client
import os
import sys
import pprint

if not os.path.exists("saved_documents"):
    os.mkdir("saved_documents")

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = True

try:
    paths = (sys.argv[1],)
except IndexError:
    paths = (path for path in os.listdir("docs") if path.endswith(".html") or path.endswith(".md"))

for file_name in paths:
    print("Parsing: %s" % file_name)

    if file_name.endswith(".html"):
        parser = HTMLParser()
    else:
        parser = MarkdownParser()

    document = word.Documents.Add()

    renderer = COMRenderer(word, document, document.ActiveWindow.Selection)

    with open(os.path.join("html", file_name), "r") as fd:
        Html = fd.read()

    renderer.render(parser.parse(Html))

    path = os.path.abspath(os.path.join("saved_documents", file_name + ".docx"))

    document.SaveAs(path)
    #document.Close()