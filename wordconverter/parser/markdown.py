from . import BaseParser
import CommonMark


class MarkdownParser(BaseParser):
    def __init__(self):
        pass

    def parse(self, content):
        p = CommonMark.DocParser()
        ast = p.parse(content)
        CommonMark.dumpAST(ast)