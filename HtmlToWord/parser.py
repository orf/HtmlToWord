from elements import *
from collections import defaultdict
import BeautifulSoup
import warnings
import functools
import itertools

#BeautifulSoup.name2codepoint["nbsp"] = ord(" ")

ElementMappings = {
    "p": Paragraph,
    "b": Bold,
    "strong": Bold,
    "br": Break,
    "div": Div,
    "em": Italic,
    "i": Italic,
    "u": UnderLine,
    "span": Span,

    "ul": UnorderedList,
    "ol": OrderedList,
    "li": ListElement,

    "table": Table,
    "tbody": TableBody,
    "thead": TableHead,
    "tr": TableRow,
    "td": TableCell,
    "th": TableCell,

    "img": Image,
    "a": HyperLink,

    "h1": Heading1,
    "h2": Heading2,
    "h3": Heading3,
    "h4": Heading4,

    "html": HTML,
    "pre": Pre,

    "blockquote": IgnoredElement,
    "wbr": IgnoredElement
}


class Parser(object):
    def __init__(self, ElementMap=None, Word=None):
        self.ElementMappings = ElementMap or ElementMappings
        self.Word = Word

        # Rendering hooks:
        # preRenderHook is ran before the element is rendered to the page.
        # Returning False will cause the element (and all children) to be ignored and not rendered
        self.preRenderHooks = defaultdict(list)
        # postRenderHook is ran after the element and its children are rendered. Its return value is ignored
        self.postRenderHooks = defaultdict(list)
        # renderHook is ran after the element has been rendered and before the children have been rendered.
        self.renderHooks = defaultdict(list)

    def AddElement(self, element, tags):
        for tag in tags:
            self.ElementMappings[tag] = element

    def ReplaceElement(self, old_element, new_element):
        for tag, element in self.ElementMappings.items():
            if element == old_element:
                self.ElementMappings[tag] = new_element

    def Parse(self, html):
        """
        I take HTML or a BeautifulSoup instance and return a list of parsed elements for use with Render.
        """

        if isinstance(html, basestring):
            old_html = html
            html = BeautifulSoup.BeautifulSoup(old_html, convertEntities="xhtml")
            if html.findChild("html") is None:
                # No HTML root tag.
                html = BeautifulSoup.BeautifulSoup("<html>%s</html>" % old_html,
                                                   convertEntities="xhtml")

        return (item for item in (self._Parse(None, child)
                                  for child in html.childGenerator())
                if item is not None)

    def _Parse(self, parent, element):
        if isinstance(element, BeautifulSoup.NavigableString):
            if element.isspace():
                return None
            return Text(element)

        ElementInstance = self.ElementMappings.get(element.name, IgnoredElement)()

        if isinstance(ElementInstance, IgnoredElement):
            warnings.warn("Element %s is ignored" % element.name)

        ElementInstance.SetAttrs(dict(element.attrs))

        if ElementInstance.IsIgnored:
            if parent is None:
                ElementInstance = self.ElementMappings["html"]()
            else:
                ElementInstance = parent

        for child in element:
            tchild = self._Parse(ElementInstance, child)
            if tchild is None:
                continue

            if ElementInstance.IsChildAllowed(tchild) and not tchild is ElementInstance:
                ElementInstance.Add(tchild)
            else:
                if ElementInstance.IsElementIgnored(tchild):
                    # Ok, its ignored. Replace tchild with an IgnoredElement and continue on our merry way
                    for child in tchild.GetChildren():
                        ElementInstance.Add(child)
                    #ElementInstance.Add(self._ConvertToIgnoredElement(tchild))

        return ElementInstance

    def _ConvertToIgnoredElement(self, element):
        new_ignored = IgnoredElement()
        for child in element.GetChildren():
            new_ignored.Add(child)
        return new_ignored

    def Render(self, Word, elements, selection, Parent=None):
        for element in elements:
            element.SetWord(Word)
            element.SetParent(Parent)
            element.SetSelection(selection)

            if self.runCallbacks(element, self.preRenderHooks, break_on_false=True) is False:
                continue

            with element as el:
                if el is not None:
                    self.runCallbacks(element, self.renderHooks)
                    self.Render(Word, el.GetChildren(), selection, Parent=el)

            self.runCallbacks(element, self.postRenderHooks)

    # Callback logic ahoy
    def AddPreRenderCallback(self, element_class, callback):
        self.preRenderHooks[element_class].append(callback)

    def AddRenderCallback(self, element_class, callback):
        self.renderHooks[element_class].append(callback)

    def AddPostRenderCallback(self, element_class, callback):
        self.postRenderHooks[element_class].append(callback)

    def runCallbacks(self, element, dct, break_on_false=False):
        #dct = self.preRenderHooks if pre is True else self.postRenderHooks

        if len(dct) == 0:
            return True

        target_classes = filter(functools.partial(isinstance, element), dct.keys())
        #    (el_class for el_class in dct.keys() if isinstance(element, el_class))
        callbacks = itertools.chain.from_iterable((dct[k] for k in target_classes))

        return self._runCallbacks(element, callbacks, break_on_false)

    def _runCallbacks(self, element, callbacks, break_on_false=False):
        for callback in callbacks:
            if callback(element) is False and break_on_false is True:
                return False
        return True

    def ParseAndRender(self, html, Word, selection):
        elements = self.Parse(html)
        self.Render(Word, elements, selection)
