from HtmlToWord.elements.base import *
from win32com.client import constants

class Break(ChildlessElement):
    pass

class Div(BaseElement):
    pass


class Image(ChildlessElement):
    def StartRender(self):
        url = self.GetAttrs()["src"]
        caption = self.GetAttrs()["alt"]
        self.Image = self.selection.InlineShapes.AddPicture(FileName=url)
        with self.With(self.Image.Borders) as Borders:
            Borders.OutsideLineStyle = constants.wdLineStyleSingle
            Borders.OutsideColor = constants.wdColorPaleBlue
        self.selection.TypeParagraph()

        if caption:
            style = self.selection.Range.Style
            self.selection.Range.Style = self.GetDocument().Styles("caption")
            self.selection.TypeText(caption)
            self.selection.Style = style

class HyperLink(BaseElement):
    def StartRender(self):
        self.start_range = self.selection.Range.End

    def EndRender(self):
        self.GetDocument().Hyperlinks.Add(Anchor=self.GetDocument().Range(Start=self.start_range,
            End=self.selection.Range.End),
            Address=self.GetAttrs()["href"])