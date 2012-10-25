from elements import *
import BeautifulSoup
import warnings

#BeautifulSoup.name2codepoint["nbsp"] = ord(" ")

ElementMappings = {
    "p":Paragraph,
    "b":Bold,
    "strong":Bold,
    "br":Break,
    "div":Div,
    "em":Italic,
    "i":Italic,
    "u":UnderLine,
    "span":Span,

    "ul":UnorderedList,
    "ol":OrderedList,
    "li":ListElement,

    "table":Table,
    "tbody":TableBody,
    "tr":TableRow,
    "td":TableCell,

    "img":Image,
    "a":HyperLink,

    "h1":Heading1,
    "h2":Heading2,
    "h3":Heading3,
    "h4":Heading4,

    "html":HTML,
    "pre":Pre,

    "blockquote":IgnoredElement,
    "wbr":IgnoredElement
}

class Parser(object):
    def __init__(self, ElementMap=None, ReplaceNewlines=True,Word=None):
        self.ElementMappings = ElementMap or ElementMappings
        self.ReplaceNewlines=ReplaceNewlines
        self.Word = Word

    def AddElement(self, element, tags):
        for tag in tags:
            self.ElementMappings[tag] = element

    def ReplaceElement(self, old_element, new_element):
        for tag,element in self.ElementMappings.items():
            if element == old_element:
                self.ElementMappings[tag] = new_element

    def Parse(self, html):
        """
        I take HTML or a BeautifulSoup instance and return a list of parsed elements for use with Render.
        """
        if self.ReplaceNewlines:
            html = html.strip("\r").strip("\n")
        if isinstance(html, basestring):
            html = BeautifulSoup.BeautifulSoup(html,convertEntities="xhtml")

        return [self._Parse(None,child) for child in html.childGenerator()]

    def _Parse(self, parent, element):
        if isinstance(element, BeautifulSoup.NavigableString):
            return Text(element)

        ElementInstance = self.ElementMappings.get(element.name, IgnoredElement)()
        if isinstance(ElementInstance, IgnoredElement):
            warnings.warn("Element %s is ignored"%element.name)

        ElementInstance.SetAttrs(dict(element.attrs))

        if ElementInstance.IsIgnored:
            if parent is None:
                ElementInstance = self.ElementMappings["html"]()
            else:
                ElementInstance = parent

        for child in element:
            tchild = self._Parse(ElementInstance, child)
            if ElementInstance.IsChildAllowed(tchild) and not tchild is ElementInstance:
                ElementInstance.Add(tchild)

        return ElementInstance

    def Render(self, Word, elements, selection, Parent=None):
        for element in elements:
            element.SetWord(Word)
            element.SetParent(Parent)
            element.SetSelection(selection)

            with element as el:
                self.Render(Word, el.GetChildren(), selection, Parent=el)

    def ParseAndRender(self, html, Word, selection):
        elements = self.Parse(html)
        self.Render(Word, elements, selection)