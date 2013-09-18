import contextlib
from collections import defaultdict
import warnings


class BaseElement(object):
    AllowedChildren = []
    IgnoredChildren = []
    IsIgnored = False

    # StripTextAfter is used by elements that contain text, such as paragraphs. Text objects after will be stripped
    # of any whitespace
    StripTextAfter = False
    # StripAfterFirstElement is used by Paragraphs. The first child element (if its text) will be stripped of whitespace.
    StripFirstElementText = False

    def __init__(self, children=None, attributes=None):
        self.children = children or []
        self.selection = None
        self.parent = None
        self.attrs = attributes or {}
        self.__shouldCallEndRender = True

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

    def IsEmpty(self):
        return len(self.GetChildren()) == 0

    def GetLastChild(self):
        if not self.children:
            return None
        else:
            return self.children[-1]

    def GetChildren(self):
        return self.children

    def GetChildByName(self, name):
        """
        Returns (idx, child) or None
        """
        for idx,child in enumerate(self.GetChildren()):
            if child.GetName() == name:
                return idx, child

        return None, None

    def GetAllowedChildren(self):
        return []  # Represents any child

    def DelegateChildrenToElement(self, new_parent):
        """
        Give me an element (with no children) and I will fill it with my current children. E.G:
        <p>text</p> -> DelegateChildren(Bold()) -> <p><strong>text</strong></p>
        """
        new_parent.SetParent(self)

        for child in self.GetChildren():
            child.SetParent(new_parent)
            new_parent.Add(child)

        self.children = [new_parent]

    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        return "<%s: Children = %s>" % (self.__class__.__name__, self.children)

    def SetSelection(self, selection):
        self.selection = selection

    def StartRender(self):
        return

    def EndRender(self):
        return

    def _StartRender(self):
        if not isinstance(self, ChildlessElement):
            # We are not a childless element, if we are empty or all our children are
            # then we don't render
            if self.IsEmpty():
                self.__shouldCallEndRender = False
                return

            o = self
            while len(o.GetChildren()) <= 1:
                # While the count of o's children is 0 or 1

                if isinstance(o, ChildlessElement):
                    break

                if o.IsEmpty():
                    self.__shouldCallEndRender = False
                    return

                o = o.GetChildren()[0]

        self.StartRender()
        return True

    def _EndRender(self):
        if self.__shouldCallEndRender:
            self.EndRender()

    def SetParent(self, parent):
        self.parent = parent

    def GetParent(self):
        el = self.parent

        while el is not None and el.IsIgnored:
            el = el.parent

        if el is None:
            warnings.warn("Element %s has no non-IgnoredElement parents and GetParent was called on it" % self)
        return el

    def GetChildIndex(self, child):
        try:
            return self.GetChildren().index(child)
        except ValueError:
            warnings.warn("Element %s is not in %s's children" % (child, self))
            return None

    @classmethod
    def GetName(cls):
        return cls.__name__

    def __enter__(self):
        if self._StartRender():
            return self

    def __exit__(self, *args, **kwargs):
        self._EndRender()
        return False


class IgnoredElement(BaseElement):
    IsIgnored = True


class ChildlessElement(BaseElement):
    def IsChildAllowed(self, child):
        return False


class HTML(BaseElement):
    pass