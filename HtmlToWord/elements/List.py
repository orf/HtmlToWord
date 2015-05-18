from HtmlToWord.elements.Base import *
from win32com.client import constants


class List(BlockElement):
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

    AllowedChildren = ["ListElement", "OrderedList", "UnorderedList"]

    def StartRender(self):
        parent = self.GetParent()
        # Here we check to see if we have a parent and if they are a listelement or a orderedlist.
        if parent is not None and parent.GetName() in self.AllowedChildren:
            # Our parent is a ListElement, this means we are a nested sub-list within a <li> tag.
            self.selection.Range.ListFormat.ListIndent()
        else:
            self.selection.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=self.GetWord().ListGalleries(self.GetTemplate()).ListTemplates(1),
                ContinuePreviousList=False,
                DefaultListBehavior=constants.wdWord10ListBehavior
            )

    def EndRender(self):
        parent = self.GetParent()

        # Before outdenting or closing the list we need to be in a newline,
        # else we will remove the last element as well. We use self.EndRender so it doesn't
        # insert multiple line breaks if the nested list is the last element of the parent.
        self.addLineBreak()

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


class ListElement(BlockElement):
    IgnoredChildren = ["Break"]
