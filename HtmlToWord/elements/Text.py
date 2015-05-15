from HtmlToWord.elements.Base import *
from HtmlToWord.elements.List import List
from win32com.client import constants
import re


class Bold(InlineElement):
    def ApplyFormatting(self, start_pos, end_pos):
        rng = super(Bold, self).ApplyFormatting(start_pos, end_pos)
        if rng:
            rng.Font.Bold = True

class Italic(InlineElement):
    def ApplyFormatting(self, start_pos, end_pos):
        rng = super(Italic, self).ApplyFormatting(start_pos, end_pos)
        if rng:
            rng.Font.Italic = True


class UnderLine(InlineElement):
    def ApplyFormatting(self, start_pos, end_pos):
        rng = super(UnderLine, self).ApplyFormatting(start_pos, end_pos)
        if rng:
            rng.Font.UnderlineColor = constants.wdColorAutomatic
            rng.Font.Underline = constants.wdUnderlineSingle


class Text(ChildlessElement):
    _COLLAPSE_REGEX = re.compile(r'\s+')

    def __init__(self, text):
        super(Text, self).__init__()
        self.Text = text

    def IsText(self):
        return True

    def GetText(self):
        parent = self.GetParent()
        txt = self.Text

        if parent is not None:
            child_index = parent.GetChildIndex(self)
            previous = parent

            if child_index != 0:
                # We are not the first child
                previous = self.GetParent().GetChildren()[child_index - 1]

            if previous.StripTextAfter or (parent.StripFirstElementText and child_index == 0):
                # We are the first child or our parent is telling us to strip :O
                txt = txt.lstrip()

            #if child_index == len(parent.GetChildren()) - 1:
                # We are the last child, strip from the right
            #    txt = txt.rstrip()

            if not isinstance(parent, Pre):
                txt = self._COLLAPSE_REGEX.sub(' ', txt)

        return txt

    def SetText(self, text):
        self.Text = text

    def StartRender(self):
        if self.Text.isspace():
            return
        self.selection.TypeText(self.GetText())

    def __repr__(self):
        return "<Text: %s>" % repr(self.Text)


class Paragraph(BlockElement):
    StripTextAfter = True

    def StartRender(self):
        if self.HasChild("Break"):
            self.PreviousStyle = self.selection.Style
            self.selection.Style = self.GetDocument().Styles("No Spacing")

    def EndRender(self):
        if self.HasChild("Break"):
            self.selection.Style = self.PreviousStyle


class Pre(BlockElement):
    StripFirstElementText = True
    PRE_FORMATTED = True

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
