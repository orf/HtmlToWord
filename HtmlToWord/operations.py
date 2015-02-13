import pprint


class Operation(object):
    requires = set()

    def __init__(self, children=None, **kwargs):
        self.children = children or []
        self.args = []

        for kwarg, val in kwargs.items():
            if kwarg not in self.requires:
                raise RuntimeError("Unexpected argument {0}".format(kwarg))
            setattr(self, kwarg, val)

    def add_child(self, child):
        self.children.append(child)

    def add_children(self, children):
        self.children.extend(children)

    def __repr__(self):
        return "<{0}: {1}>".format(self.__class__.__name__, self.children)


class ChildlessOperation(Operation):
    def __init__(self):
        super().__init__([])

    def __repr__(self):
        return "<{0}>".format(self.__class__.__name__)


class IgnoredElement(Operation):
    pass


class Group(Operation):
    pass


class Bold(Operation):
    pass


class Italic(Operation):
    pass


class UnderLine(Operation):
    pass


class Text(ChildlessOperation):
    def __init__(self, text):
        self.text = text
        super().__init__()


class Paragraph(Operation):
    pass


class BlockParagraph(Operation):
    """
    Same as Paragraph but doesn't add a newline at the end
    """


class CodeBlock(Operation):
    pass


class Font(Operation):
    requires = {"name"}


class Image(Operation):
    requires = {"src"}


class HyperLink(Operation):
    requires = {"href"}