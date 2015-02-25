from . import BaseParser
from ..operations import Paragraph, Bold, Italic, UnderLine, Text,\
    CodeBlock, Group, IgnoredOperation, Style, Image, HyperLink, BulletList,\
    NumberedList, ListElement, BaseList, Table, TableRow, TableCell, TableHeading
import bs4
from functools import partial
import cssutils


class HTMLParser(BaseParser):
    def __init__(self):
        # Preserve whitespace as-is
        self.preserve_whitespace = {
            CodeBlock
        }
        # Strip whitespace but keep spaces between tags
        self.respect_whitespace = {
            Paragraph, Bold, Italic, UnderLine
        }

        self.mapping = {
            "p": Paragraph,
            "b": Bold,
            "strong": Bold,
            "i": Italic,
            "em": Italic,
            "u": UnderLine,
            "pre": CodeBlock,
            "div": Group,

            "h1": partial(Style, name="Heading 1"),
            "h2": partial(Style, name="Heading 2"),
            "h3": partial(Style, name="Heading 3"),
            "h4": partial(Style, name="Heading 4"),

            "ul": BulletList,
            "ol": NumberedList,
            "li": ListElement,

            "img": Image,
            "a": HyperLink,
            "html": Group,

            "table": Table,
            "tr": TableRow,
            "td": TableCell,
            "th": TableHeading,
        }

    def parse(self, content):
        parser = bs4.BeautifulSoup(content)

        tokens = []

        for element in parser.childGenerator():
            item = self.build_element(element)

            if item is None:
                continue

            tokens.append(item)
        import pprint
        pprint.pprint(tokens)
        return tokens

    def build_element(self, element, whitespace="ignore"):
        if isinstance(element, bs4.Comment):
            return None

        if isinstance(element, bs4.NavigableString):
            if element.isspace():
                if whitespace == "preserve":
                    return Text(text=str(element))
                elif whitespace == "ignore":
                    return None
                elif whitespace == "respect":
                    if isinstance(element.previous_sibling, bs4.NavigableString):
                        return None
                    return Text(text=" ")

            return Text(text=str(element.strip()))

        cls = self.mapping.get(element.name, IgnoredOperation)

        if cls is Image:
            cls = partial(Image,
                          height=element.attrs.get("height", None),
                          width=element.attrs.get("width", None),
                          caption=element.attrs.get("alt", None),
                          location=element.attrs["src"])
        elif cls is HyperLink:
            cls = partial(HyperLink, location=element.attrs["href"])

        instance = cls()

        if cls in self.respect_whitespace:
            whitespace = "respect"
        elif cls in self.preserve_whitespace:
            whitespace = "preserve"

        for child in element.childGenerator():
            item = self.build_element(child, whitespace=whitespace)
            if item is None:
                continue

            if isinstance(instance, BaseList) and not isinstance(item, ListElement):
                # Wrap the item in a ListElement
                item = ListElement(children=[item])

            if isinstance(item, IgnoredOperation):
                instance.add_children(item.children)
            else:
                instance.add_child(item)

        return instance