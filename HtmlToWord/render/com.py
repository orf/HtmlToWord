from . import Renderer, renders
from ..operations import HyperLink, Text


class COMRenderer(Renderer):
    def __init__(self, word, selection):
        super().__init__()

    @renders(HyperLink)
    def hyperlink(self, op):
        print(op.href)
        yield

    @renders(Text)
    def text(self, op):
        print(op.text)