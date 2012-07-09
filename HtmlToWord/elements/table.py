from HtmlToWord.elements.base import *
from win32com.client import constants

class Table(BaseElement):
    """
    Ding dong i'm a f*****g table.
    I'm not a very good table, I can't have merged cells, and this makes me sad.
    """
    AllowedChildren = ["TableRow","TableBody"]

    def __init__(self, *args, **kwargs):
        super(Table, self).__init__(*args, **kwargs)
        self.TableRow = 0

    def GetDimentions(self):
        if self.GetChildren()[0].GetName() == "TableBody":
            rows = self.GetChildren()[0].CountRows()
        else:
            rows = len(self.GetChildren())
        cells = len(self.GetChildren()[0].GetChildren())

        return rows,cells

    def StartRender(self):
        rng = self.selection.Range
        self.selection.TypeParagraph()
        self._end_range = self.selection.Range

        rows,cells = self.GetDimentions()

        self.Table = self.selection.Tables.Add(rng,
            NumRows=rows,
            NumColumns=cells,
            AutoFitBehavior=constants.wdAutoFitFixed,
        )

        for count,child in enumerate(self.GetChildren()):
            child.SetRow(self.Table.Rows(count+1))

        self.Table.Style = "Table Grid"
        self.Table.AllowAutoFit = True

    def EndRender(self):
        self.Table.Columns.AutoFit()
        self._end_range.Select()

class TableBody(IgnoredElement):
    pass

class TableRow(BaseElement):
    AllowedChildren = ["TableCell"]

    def __init__(self, *args, **kwargs):
        super(TableRow, self).__init__(*args, **kwargs)
        self.Row = None

    def SetRow(self, Row):
        self.Row = Row

    def StartRender(self):

        for count,child in enumerate(self.GetChildren()):
            assert child.GetName() == "TableCell", "Child of TableRow is not TableCell! Its %s"%child.GetName()
            child.SetCell(self.Row.Cells(count+1))


class TableCell(BaseElement):

    def SetCell(self, Cell):
        self.Cell = Cell

    def StartRender(self):
        self.Cell.Range.Select()
