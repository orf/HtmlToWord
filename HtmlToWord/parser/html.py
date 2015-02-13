from . import BaseParser
from ..operations import Paragraph, Bold, Italic, UnderLine, Text,\
    CodeBlock, Group, IgnoredElement, Font, Image, HyperLink
import bs4
from functools import partial


class HTMLParser(BaseParser):
    def __init__(self):
        self.mapping = {
            "p": Paragraph,
            "b": Bold,
            "strong": Bold,
            "i": Italic,
            "em": Italic,
            "u": UnderLine,
            "pre": CodeBlock,
            "div": Group,

            "h1": partial(Font, name="Heading 1"),
            "h2": partial(Font, name="Heading 2"),
            "h3": partial(Font, name="Heading 3"),
            "h4": partial(Font, name="Heading 4"),

            "img": Image,
            "a": HyperLink,
            "html": Group
        }

    def parse(self, content):
        parser = bs4.BeautifulSoup(content)

        tokens = []

        for element in parser.childGenerator():
            item = self._build(element)

            if item is None:
                continue

            tokens.append(item)

        return tokens

    def _build(self, element):
        if isinstance(element, bs4.NavigableString):
            if element.isspace():
                return None
            return Text(str(element))

        cls = self.mapping.get(element.name, IgnoredElement)

        if cls is Image:
            cls = partial(HyperLink, location=element.attrs["src"])
        elif cls is HyperLink:
            cls = partial(HyperLink, href=element.attrs["href"])

        instance = cls()

        for child in element.childGenerator():
            item = self._build(child)
            if item is None:
                continue
            instance.add_child(item)

        return instance