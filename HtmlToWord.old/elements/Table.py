from HtmlToWord.elements.Base import *
from win32com.client import constants


class Table(BaseElement):
    """
    Ding dong i'm a f*****g table.
    I'm not a very good table, I can't have merged cells, and this makes me sad.
    I support two types of table:
     * Simple ones with no tbody or thead (just <table><tr><td>...)
     * Complex ones with a tbody and thead element.
    """
    AllowedChildren = ["TableRow", "TableBody", "TableHead"]

    def __init__(self, *args, **kwargs):
        super(Table, self).__init__(*args, **kwargs)
        self.TableRow = 0
        self.HasHeader = False

    def GetDimentions(self):
        rows, columns = 0, 0

        if self.HasChild("TableHead"):
            rows += 1

        if self.HasChild("TableBody"):
            idx, tbody = self.GetChildByName("TableBody")
            rows += len(tbody.GetChildren())
            columns = len(tbody.GetChildren()[0].GetChildren())
        else:
            rows = len(self.GetChildren())
            idx, first_row = self.GetChildByName("TableRow")
            columns = len(first_row.GetChildren())

        return rows, columns

    def StartRender(self):
        self.HasHeader = self.GetChildByName("TableHead")[0] is not None

        rng = self.selection.Range
        self.selection.TypeParagraph()
        self._end_range = self.selection.Range

        rows, cells = self.GetDimentions()
        self.Table = self.selection.Tables.Add(rng,
                                               NumRows=rows,
                                               NumColumns=cells,
                                               AutoFitBehavior=constants.wdAutoFitFixed)

        for count, child in enumerate(self.GetChildren()):
            child.SetRow(self.Table.Rows(count+1))

        self.Table.Style = "Table Grid"
        self.Table.AllowAutoFit = True

    def EndRender(self):
        self.Table.Columns.AutoFit()
        self._end_range.Select()


class TableBody(IgnoredElement):
    pass


class TableHead(BaseElement):
    def SetRow(self, Row):
        self.GetChildren()[0].SetRow(Row)


class TableRow(BaseElement):
    AllowedChildren = ["TableCell"]

    def __init__(self, *args, **kwargs):
        super(TableRow, self).__init__(*args, **kwargs)
        self.Row = None
        self.IsHeader = False

    def SetRow(self, Row):
        self.Row = Row

    def StartRender(self):
        self.IsHeader = self.GetParent().GetName == "TableHead"
        for count, child in enumerate(self.GetChildren()):
            assert child.GetName() == "TableCell", "Child of TableRow is not TableCell! Its %s" % child.GetName()
            child.SetCell(self.Row.Cells(count+1))


class TableCell(BaseElement):
    StripTextAfter = True

    def SetCell(self, Cell):
        self.Cell = Cell

    def StartRender(self):
        self.Cell.Range.Select()
