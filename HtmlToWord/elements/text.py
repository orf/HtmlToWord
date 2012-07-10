from HtmlToWord.elements.base import *
from win32com.client import constants

class Bold(BaseElement):
    def StartRender(self):
        self.selection.BoldRun()

    def EndRender(self):
        self.selection.BoldRun()

class Italic(BaseElement):
    def StartRender(self):
        self.selection.ItalicRun()

    def EndRender(self):
        self.selection.ItalicRun()

class UnderLine(BaseElement):
    def StartRender(self):
        with self.With(self.selection.Font) as Font:
            Font.UnderlineColor = constants.wdColorAutomatic
            Font.Underline = constants.wdUnderlineSingle

    def EndRender(self):
        with self.With(self.selection.Font) as Font:
            Font.UnderlineColor = constants.wdColorAutomatic
            Font.Underline = constants.wdUnderlineNone

class Text(BaseElement):
    def __init__(self, text):
        super(Text, self).__init__()
        self.Text = text

    def IsText(self):
        return True

    def GetText(self):
        return self.Text

    def StartRender(self):
        if self.GetText() == "\n":
            return
        self.selection.TypeText(self.GetText())

    def __repr__(self):
        return "<Text: %s>"%repr(self.Text)

class Paragraph(BaseElement):
    def EndRender(self):
        self.selection.TypeParagraph()