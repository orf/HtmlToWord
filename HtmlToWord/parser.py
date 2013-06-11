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
    def __init__(self, ElementMap=None, ReplaceNewlines=True, Word=None,
                 preRenderHook=None, postRenderHook=None):
        self.ElementMappings = ElementMap or ElementMappings
        self.ReplaceNewlines = ReplaceNewlines
        self.Word = Word

        # Rendering hooks:
        # preRenderHook is ran before the element is rendered to the page.
        # Returning False will cause the element (and all children) to be ignored and not rendered
        self.preRenderHooks = defaultdict(list)
        # postRenderHook is ran after the element is rendered. Its return value is ignored
        self.postRenderHooks = defaultdict(list)

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
        if self.ReplaceNewlines:
            html = html.replace("\n", "").replace("\r", "")
        if isinstance(html, basestring):
            html = BeautifulSoup.BeautifulSoup(html, convertEntities="xhtml")

        return (self._Parse(None, child) for child in html.childGenerator())

    def _Parse(self, parent, element):
        if isinstance(element, BeautifulSoup.NavigableString):
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
            if ElementInstance.IsChildAllowed(tchild) and not tchild is ElementInstance:
                ElementInstance.Add(tchild)
            else:
                if ElementInstance.IsElementIgnored(tchild):
                    # Ok, its ignored. Replace tchild with an IgnoredElement and continue on our merry way
                    ElementInstance.Add(self._ConvertToIgnoredElement(tchild))

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

            if self.runCallbacks(element, break_on_false=True) is False:
                continue

            with element as el:
                self.Render(Word, el.GetChildren(), selection, Parent=el)

            self.runCallbacks(element, pre=False)

    # Callback logic ahoy
    def AddPreRenderCallback(self, element_class, callback):
        self.preRenderHooks[element_class].append(callback)

    def AddPostRenderCallback(self, element_class, callback):
        self.postRenderHooks[element_class].append(callback)

    def runCallbacks(self, element, pre=True, break_on_false=False):
        dct = self.preRenderHooks if pre is True else self.postRenderHooks

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
