from wordinserter import insert, parse
from comtypes.client import CreateObject
import os
import sys
import time

if not os.path.exists("saved_documents"):
    os.mkdir("saved_documents")

word = CreateObject("Word.Application")
word.Visible = True

from comtypes.gen import Word as constants

try:
    paths = (sys.argv[1],)
except IndexError:
    paths = (path for path in os.listdir("docs") if path.endswith(".html") or path.endswith(".md"))

for file_name in paths:
    print("Parsing: %s" % file_name)

    document = word.Documents.Add()

    with open(os.path.join("docs", file_name), "r") as fd:
        text = fd.read()

    start = time.time()
    operations = parse(text, "html" if file_name.endswith(".html") else "markdown")
    print(time.time() - start)

    render_start = time.time()

    insert(operations, document=document, constants=constants)
    print(time.time() - render_start)



    path = os.path.abspath(os.path.join("saved_documents", file_name + ".docx"))

    #document.SaveAs(path)
    #document.Close()