import contextlib
from collections import defaultdict
import warnings


class BaseElement(object):
    AllowedChildren = []
    IgnoredChildren = []
    IsIgnored = False

    def __init__(self, children=None, attributes=None):
        self.children = children or []
        self.selection = None
        self.parent = None
        self.attrs = attributes or {}

    @contextlib.contextmanager
    def With(self, item):
        """ Convenience method. Not sure why its here """
        yield item

    def IsChildAllowed(self, child):
        assert not (self.AllowedChildren and self.IgnoredChildren), "Only AllowedChildren OR IgnoredChildren allowed"

        if not self.AllowedChildren and not self.IgnoredChildren:
            return True

        if self.IgnoredChildren:
            return not child.GetName() in self.IgnoredChildren
        else:
            return child.GetName() in self.AllowedChildren

    def IsElementIgnored(self, element):
        return element.GetName() in self.IgnoredChildren

    def SetAttrs(self, attrs):
        self.attrs = defaultdict(lambda: None, attrs)

    def GetAttrs(self):
        return self.attrs

    def SetWord(self, word):
        self.word = word
        self.document = word.ActiveDocument

    def GetDocument(self):
        return self.document

    def GetWord(self):
        return self.word

    def Add(self, child):
        self.children.append(child)

    def HasChild(self, child):
        if not isinstance(child, basestring):
            child = child.GetName()

        for c in self.GetChildren():
            if child == c.GetName():
                return True

    def IsText(self):
        return False

    def GetLastChild(self):
        if not self.children:
            return None
        else:
            return self.children[-1]

    def GetChildren(self):
        return self.children

    def GetAllowedChildren(self):
        return []  # Represents any child

    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        return "<%s: Children = %s>"%(self.__class__.__name__, self.children)

    def SetSelection(self, selection):
        self.selection = selection

    def StartRender(self):
        return

    def EndRender(self):
        return

    def SetParent(self, parent):
        self.parent = parent

    def GetParent(self):
        el = self.parent

        while el is not None and el.IsIgnored:
            el = el.parent

        if el is None:
            warnings.warn("Element %s has no non-IgnoredElement parents and GetParent was called on it" % self)
        return el

    @classmethod
    def GetName(cls):
        return cls.__name__

    def __enter__(self):
        self.StartRender()
        return self

    def __exit__(self, *args, **kwargs):
        self.EndRender()
        return False


class IgnoredElement(BaseElement):
    IsIgnored = True


class ChildlessElement(BaseElement):
    def IsChildAllowed(self, child):
        return False


class HTML(BaseElement):
    pass