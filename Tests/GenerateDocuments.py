import HtmlToWord
import win32com.client
import os
import sys
import pprint

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
    print "-" * 30
    print "Parsing: %s" % file_name

    document = word.Documents.Add()

    with open(os.path.join("html", file_name), "r") as fd:
        Html = fd.read()
    print "Parsed HTML:"
    pprint.pprint(list(parser.Parse(Html)))
    print "Rendering..."

    #parser.preRenderHook = lambda el: sys.stdout.write("preElement %s\n" % el)
    #parser.postRenderHook = lambda el: sys.stdout.write("postRender: %s\n" % el)
    from HtmlToWord.elements.Table import Table
    from HtmlToWord.elements.Base import BaseElement

    def _postRenderHook(element):
        ''' Make all tables a blue style'''

        print "Element %s" % element
        if isinstance(element, Table):
            # Styles: http://msdn.microsoft.com/en-us/library/office/ff835210(v=office.14).aspx
            if element.HasHeader:
                element.Table.Style = -178

    parser.AddPostRenderCallback(BaseElement, _postRenderHook)

    parser.ParseAndRender(Html, word, document.ActiveWindow.Selection)
    path = os.path.abspath(os.path.join("saved_documents", file_name + ".docx"))

    document.SaveAs(path)
    document.Close()

    print "-" * 30