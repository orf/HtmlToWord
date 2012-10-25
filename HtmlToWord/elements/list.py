from HtmlToWord.elements.base import *
from win32com.client import constants

class List(BaseElement):
    """
    I am a list, I can be ordered or unordered and I can have sub lists within me.
    There are two common ways to express sub lists using HTML and this class handles both of them:

    <ul>
        <li>Some text here
            <ul>
                <li>My parent ul tag is nested within a li tag</li>
            </ul>
        </li>
    </ul>

    <ul>
        <li>Some text here</li>
        <ul>
            <li>My parent ul tag is *not* nested within a li tag</li>
        </ul>
    </ul>


    """

    AllowedChildren = ["ListElement", "OrderedList","UnorderedList"]

    def StartRender(self):
        parent = self.GetParent()
        # Here we check to see if we have a parent and if they are a listelement or a orderedlist.
        if parent is not None and parent.GetName() in self.AllowedChildren:
            if parent.GetName() == "ListElement":
                # Our parent is a ListElement, this means we are a nested sub-list within a <li> tag.
                # We make a new paragraph here because there will be text above us (not a <p> element)
                # so we need to make a new list entry below, which will then be indented.
                self.selection.TypeParagraph()
            self.selection.Range.ListFormat.ListIndent()
        else:
            self.selection.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=self.GetWord().ListGalleries(self.GetTemplate()).ListTemplates(1),
                ContinuePreviousList=False,
                DefaultListBehavior=constants.wdWord10ListBehavior
            )
        pass

    def EndRender(self):
        parent = self.GetParent()
        if parent is not None and parent.GetName() in self.AllowedChildren:
            # If we are nested then we just outdent the list
            self.selection.Range.ListFormat.ListOutdent()
        else:
            # Otherwise we clear the list entirely.
            self.selection.Range.ListFormat.RemoveNumbers(NumberType=constants.wdNumberParagraph)

    def GetTemplate(self):
        """ Return the template type. Override me in a subclass please.
        I have to be a function because the constants.* pseudo-module is not
        populated until word is started.
        """
        return None

class OrderedList(List):
    def GetTemplate(self):
        return constants.wdNumberGallery

class UnorderedList(List):
    def GetTemplate(self):
        return constants.wdBulletGallery

class ListElement(BaseElement):
    IgnoredChildren = ["Break"]
    def EndRender(self):
        # A bit of a hack but whatever. If the last child is a OrderedList or UnorderedList
        # then we have a nested sub-list. Don't start a new paragraph because the List will
        # do this for us.
        last_child = ""
        if len(self.GetChildren()):
            last_child = self.GetChildren()[-1].GetName()

        if not last_child in ("OrderedList","UnorderedList"):
            self.selection.TypeParagraph()