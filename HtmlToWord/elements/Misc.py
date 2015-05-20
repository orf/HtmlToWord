import os
import tempfile
import urllib2
import warnings

from HtmlToWord.elements.Base import *
from win32com.client import constants
from HtmlToWord import groups


class Break(ChildlessElement):
    """
    I'm a really annoying element who gets in the way of things. I sometimes have an effect and I sometimes don't.
    If I'm in a Paragraph I cause the paragraph to change its style
    If I'm in a ListElement I mess things up so I am excluded from the party :(

    I still have a few bugs to do with me being nested in tags and then those tags being in paragraphs and lists,
    but screw it.
    """
    StripTextAfter = True

    def EndRender(self):
        self.selection.TypeParagraph()


class Div(BlockElement):
    pass


class Span(InlineElement):
    pass


class Image(ChildlessElement):
    def StartRender(self):
        url = self.GetAttrs()["src"]
        caption = self.GetAttrs()["alt"]
        height = self.GetAttrs()["height"]
        width = self.GetAttrs()["width"]
        if url.startswith('https'):
            # workaround to fetch images from https urls: in some cases MS Word is not able to correctly
            # fetch remote files over HTTPS connections, so it's worth to fetch them separately and store
            # them in a tempoarary file.
            try:
                response = urllib2.urlopen(url)
            except urllib2.URLError:
                warnings.warn('Unable to load image {url}, skipping'.format(url=url))
                return
            else:
                with tempfile.NamedTemporaryFile(delete=False) as temporary_file:
                    temporary_file.write(response.read())
                self.Image = self.selection.InlineShapes.AddPicture(FileName=temporary_file.name)
                os.remove(temporary_file.name)
        else:
            self.Image = self.selection.InlineShapes.AddPicture(FileName=url)
        if height:
            self.Image.Height = height
        if width:
            self.Image.Width = width
        if caption:
            style = self.selection.Range.Style
            self.selection.Range.Style = self.GetDocument().Styles("caption")
            self.selection.TypeText(caption)
            self.selection.TypeParagraph()


class HyperLink(InlineElement):
    # Formatting tags don't work well inside hyperlinks. Ignore them.
    IgnoredChildren = groups.FORMAT_TAGS

    def StartRender(self):
        self.start_range = self.selection.Range.End

    def EndRender(self):
        href = self.GetAttrs()["href"]
        if href:
            document_range = self.GetDocument().Range(Start=self.start_range,
                                                      End=self.selection.Range.End)

            self.GetDocument().Hyperlinks.Add(Anchor=document_range, Address=href)
