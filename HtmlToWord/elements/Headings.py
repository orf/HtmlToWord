from HtmlToWord.elements.Base import *


class BaseHeading(BlockElement):
    StyleName = None

    def StartRender(self):
        self.selection.Style = self.GetDocument().Styles(self.StyleName)


class Heading1(BaseHeading):
    StyleName = "Heading 1"


class Heading2(BaseHeading):
    StyleName = "Heading 2"


class Heading3(BaseHeading):
    StyleName = "Heading 3"


class Heading4(BaseHeading):
    StyleName = "Heading 4"