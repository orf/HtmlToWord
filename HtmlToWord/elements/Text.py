from HtmlToWord.elements.Base import *
from HtmlToWord.elements.List import List
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
        child_index = self.GetParent().GetChildIndex(self)
        if child_index == 0:
            previous = self.GetParent()
        else:
            previous = self.GetParent().GetChildren()[child_index - 1]

        txt = self.Text

        if previous.StripTextAfter or (self.GetParent().StripFirstElementText and child_index == 0):
            txt = txt.lstrip()

        if child_index == len(self.GetParent().GetChildren()) - 1:
            txt = txt.rstrip()

        return txt

    def SetText(self, text):
        self.Text = text

    def StartRender(self):
        if self.Text.isspace():
            return
        self.selection.TypeText(self.GetText())

    def __repr__(self):
        return "<Text: %s>" % repr(self.Text)


class Paragraph(BaseElement):
    StripTextAfter = True

    def StartRender(self):
        if self.HasChild("Break"):
            self.PreviousStyle = self.selection.Style
            self.selection.Style = self.GetDocument().Styles("No Spacing")

    def EndRender(self):
        if self.HasChild("Break"):
            self.selection.Style = self.PreviousStyle
        # Adding a paragprah after this looks weird as the list does this itself.
        if not isinstance(self.GetLastChild(), List):
            self.selection.TypeParagraph()


class Pre(BaseElement):
    StripFirstElementText = True

    def StartRender(self):
        self.PreviousStyle = self.selection.Style
        self.PreviousFont = self.selection.Font.Name

        self.selection.Style = self.GetDocument().Styles("No Spacing")
        self.selection.Font.Name = "Courier New"
        self.selection.Font.Size = 7

    def EndRender(self):
        self.selection.ParagraphFormat.LineSpacingRule = constants.wdLineSpace1pt5
        self.selection.TypeParagraph()
        self.selection.Style = self.PreviousStyle
        self.selection.Font.Name = self.PreviousFont
