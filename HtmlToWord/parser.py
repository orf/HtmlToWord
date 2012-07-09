from elements import *
import BeautifulSoup

BeautifulSoup.name2codepoint["nbsp"] = ord(" ")

ElementMappings = {
    "p":Paragraph,
    "b":Bold,
    "strong":Bold,
    "br":Break,
    "div":Div,
    "em":Italic,
    "i":Italic,
    "u":UnderLine,

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
}

class Parser(object):
    def __init__(self, ElementMap=None):
        self.ElementMappings = ElementMap or ElementMappings

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

        if isinstance(html, basestring):
            html = BeautifulSoup.BeautifulSoup(html,convertEntities="html")
        return [self._Parse(None,child) for child in html.childGenerator()]

    def _Parse(self, parent, element):
        if isinstance(element, BeautifulSoup.NavigableString):
            return Text(str(element))
        ElementInstance = self.ElementMappings[element.name]()
        ElementInstance.SetAttrs(dict(element.attrs))

        if ElementInstance.IsIgnored:
            ElementInstance = parent

        for child in element:
            tchild = self._Parse(ElementInstance, child)
            if ElementInstance.IsChildAllowed(tchild):
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