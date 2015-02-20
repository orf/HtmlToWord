import pprint


class Operation(object):
    requires = set()
    optional = set()

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

    def has_child(self, child_class):
        return any(isinstance(c, child_class) for c in self.children)

    def __repr__(self):
        return "<{0}: {1}>".format(self.__class__.__name__, self.children)


class ChildlessOperation(Operation):
    def __init__(self, **kwargs):
        kwargs["children"] = []
        super().__init__(**kwargs)

    def __repr__(self):
        return "<{0}>".format(self.__class__.__name__)

    def has_child(self, child_class):
        return False


class IgnoredOperation(Operation):
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
    requires = {"text"}


class Paragraph(Operation):
    pass


class BlockParagraph(Operation):
    """
    Same as Paragraph but doesn't add a newline at the end
    """


class CodeBlock(Operation):
    pass


class LineBreak(ChildlessOperation):
    pass


class Style(Operation):
    requires = {"name"}


class Font(Operation):
    optional = {"size", "color"}


class Image(ChildlessOperation):
    requires = {"location"}


class HyperLink(Operation):
    requires = {"location"}


class List(Operation):
    pass


class BulletList(List):
    pass


class NumberedList(List):
    pass


class ListElement(Operation):
    pass