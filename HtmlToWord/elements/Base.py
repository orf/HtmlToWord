import contextlib
from collections import defaultdict
import warnings


def getWdColorFromRGB(rgbAttr):
    """
    receive an rgb color attribute string like 'rgb(149, 55, 52)' and tranform it to a numeric constant
    in order to use it as a Selection.Font.Color attribute (as an item of WdColor enumeration)
    """
    try:
        values = rgbAttr[rgbAttr.find('(')+1:rgbAttr.find(')')].split(',')
    except:
        warnings.warn("getWdColorFromRGB: not possible to parse the RGB string '%s' " % rgbAttr)
        return None
    else:
        rgbstrlst = [v.strip() for v in values]
        return (int(rgbstrlst[0]) + 0x100 * int(rgbstrlst[1]) + 0x10000 * int(rgbstrlst[2]))


def getPointsFromPx(px_str):
    """
    receive an string representing the font-size attribute value in px (e.g. '16px') and tranform it
    to the equivalent value in points
    """
    try:
        px = px_str.split('px')[0]
        return int(px)*0.75
    except ValueError, IndexError:
        warnings.warn("Unable to tranform the value '%s' points" % px_str)
        return None


class BaseElement(object):
    AllowedChildren = []
    IgnoredChildren = []
    IsIgnored = False
    soup = None

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
        """
        Convenience method. If you need to access a nested object like something.somethingelse.someitem several times
        then a COM request will be issued for each nested object each time you access it. Use this instead of 
        temp = something.somethingelse.someitem
        temp.method()
        temp.method2()
        """
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

    def ApplyFormatting(self, start_pos, end_pos):
        rng = self.document.Range(start_pos, end_pos)
        for attribute, value in self.attrs.items():
            try:
                if attribute=='class' and value:
                    try:
                        rng.Style=value
                    except:
                        warnings.warn("Unable to apply the class '%s'" % (value, ))
                if attribute == 'style':
                    styles=[[s.strip() for s in x.split(':')] for x in value.split(';') if x != ""]
                    for style, val in styles:
                        if style=='font-size':
                            fontsize = getPointsFromPx(val)
                            if fontsize:
                                rng.Font.Size = fontsize
                        elif style=="color":
                            color = getWdColorFromRGB(val)
                            if color:
                                rng.Font.Color = color
                        else:
                            warnings.warn("Unable to process the style '%s' with value '%s'" % (style, val))
            except Exception as e:
                warnings.warn("Error in applying formatting - %s" % e.message)

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

        self.start_pos = self.document.ActiveWindow.Selection.Start
        self.StartRender()
        return True

    def _EndRender(self):
        if self.__shouldCallEndRender:
            if hasattr(self, 'Cell'):
                start_pos = self.Cell.Range.Start
                end_pos = self.Cell.Range.End
            else:
                start_pos = self.start_pos
                end_pos = self.document.ActiveWindow.Selection.End

            self.ApplyFormatting(start_pos, end_pos)
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
