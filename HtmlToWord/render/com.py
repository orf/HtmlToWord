from . import Renderer, renders
from ..operations import Text, Bold, Italic, UnderLine, Paragraph, LineBreak, CodeBlock, Style, Image, HyperLink, \
    BulletList, NumberedList, ListElement, List
from win32com.client import constants


class COMRenderer(Renderer):
    def __init__(self, word, document, selection):
        self.word = word
        self.document = document
        self.selection = selection
        super().__init__()

    @renders(Style)
    def style(self, op):
        old_style = self.selection.Style
        self.selection.Style = self.document.Styles(op.name)
        yield
        self.selection.TypeParagraph()
        #self.selection.Collapse(Direction=constants.wdCollapseEnd)
        #self.selection.Style = old_style

    @renders(Bold)
    def bold(self, op):
        self.selection.BoldRun()
        yield
        self.selection.BoldRun()

    @renders(Italic)
    def italic(self, op):
        self.selection.ItalicRun()
        yield
        self.selection.ItalicRun()

    @renders(UnderLine)
    def underline(self, op):
        self.selection.Font.Underline = constants.wdUnderlineSingle
        yield
        self.selection.Font.Underline = constants.wdUnderlineNone

    @renders(Text)
    def text(self, op):
        self.selection.TypeText(op.text)

    @renders(LineBreak)
    def linebreak(self, op):
        self.selection.TypeParagraph()

    @renders(Paragraph)
    def paragraph(self, op):
        previous_style = None
        if op.has_child(LineBreak):
            previous_style = self.selection.Style
            self.selection.Style = self.document.Styles("No Spacing")

        yield

        if previous_style is not None:
            self.selection.Style = previous_style

        try:
            # If the last child is a list or image then we don't want to insert a paragraph
            last_child = op.children[-1]
            if not isinstance(last_child, (List, Image)):
                self.selection.TypeParagraph()
        except IndexError:  # No children
            pass

        self.selection.TypeParagraph()

    @renders(CodeBlock)
    def pre(self, op):
        previous_style = self.selection.Style
        previous_font = self.selection.Font.Name
        self.selection.Style = self.document.Styles("No Spacing")
        self.selection.Font.Name = "Courier New"
        self.selection.Font.Size = 7

        yield

        self.selection.ParagraphFormat.LineSpacingRule = constants.wdLineSpace1pt5
        self.selection.TypeParagraph()
        self.selection.Style = previous_style
        self.selection.Font.Name = previous_font

    @renders(Image)
    def image(self, op):
        image = self.selection.InlineShapes.AddPicture(FileName=op.location)
        self.selection.TypeParagraph()

    @renders(HyperLink)
    def hyperlink(self, op):
        start_range = self.selection.Range.End
        yield
        # Inserting a hyperlink that contains different styles can reset the style. IE:
        # Link<Bold<Text>> Text
        # Bold will turn bold off, but link will reset it meaning the second Text is bold.
        # Here we just reset the style after making the hyperlink.
        style = self.selection.Style
        rng = self.document.Range(Start=start_range, End=self.selection.Range.End)
        self.document.Hyperlinks.Add(Anchor=rng, Address=op.location)
        self.selection.Collapse(Direction=constants.wdCollapseEnd)
        self.selection.Style = style

    @renders(BulletList, NumberedList)
    def render_list(self, op):
        first_list = self.selection.Range.ListFormat.ListTemplate is None

        if first_list:
            self.selection.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=self.word.ListGalleries(
                    constants.wdNumberGallery if isinstance(op, NumberedList) else constants.wdBulletGallery
                ).ListTemplates(1),
                ContinuePreviousList=False,
                DefaultListBehavior=constants.wdWord10ListBehavior
            )
        else:
            self.selection.Range.ListFormat.ListIndent()

        yield

        if first_list:
            self.selection.Range.ListFormat.RemoveNumbers(NumberType=constants.wdNumberParagraph)
        else:
            self.selection.Range.ListFormat.ListOutdent()

    @renders(ListElement)
    def list_element(self, op):
        yield
        self.selection.TypeParagraph()