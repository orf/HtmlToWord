from collections import OrderedDict

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
    CellsToMerge = []
    rows = 0
    columns = 0

    def __init__(self, *args, **kwargs):
        super(Table, self).__init__(*args, **kwargs)
        self.TableRow = 0
        self.HasHeader = False
        self.mapper = None

    def GetDimentions(self):
        rows, columns = 0, 0

        if self.HasChild("TableHead"):
            rows += 1

        if self.HasChild("TableBody"):
            idx, tbody = self.GetChildByName("TableBody")
            rows += self._GetRows(tbody.GetChildren())
            columns = self._GetColumns(tbody.GetChildren())
        else:
            rows += self._GetRows(self.GetChildren())
            columns = self._GetColumns(self.GetChildren())

        self.rows = rows
        self.columns = columns
        return rows, columns

    def _GetColumns(self, tableRows):
        max_cols = 0
        return max(
            [sum([cell.GetColspan() if cell.HasColspan() else 1 for cell in row.GetChildren()]) for row in tableRows]
        )

    def _GetRows(self, tableRows):
        max_rowspan = max(
            [max([cell.GetRowspan() if cell.HasRowspan() else 1 for cell in row.GetChildren()]) for row in tableRows]
        )
        return max_rowspan if max_rowspan > len(tableRows) else len(tableRows)

    def _MergeCells(self, alignParagraphCenter=True):
        """
        merge cells in order to deal with 'colspan' attribute.
        Cells can be adjacent or not, and can span over multiple lines.
        e.g given the following table:
        -------------------
         | c11 | c12 | c13 |
         -------------------
         | c21 | c22 | c23 |
         -------------------
         if originCell is c11 and targetCell is c23, _MergeCells() will produce
         the following result:
         -------------------
         |                 |
         |       c11       |
         |                 |
         -------------------
         Needs to be done here, after table has been rendered (otherwise the merged
         cell will have the size of the original cell), and in reversed order (from bottom/right to top/left)
         to work (after merging, the word Cell object's RowIndex and ColumnIndex properties values may change)
         """
        for originCellCoordinates, targetCellCoordinates in reversed(self.mapper.cells_to_merge):
            originCell = self.Table.Cell(*originCellCoordinates)
            targetCell = self.Table.Cell(*targetCellCoordinates)
            originCell.Merge(targetCell)
            if alignParagraphCenter:
                originCell.Range.ParagraphFormat.Alignment = constants.wdAlignParagraphCenter


    def StartRender(self):
        self.HasHeader = self.GetChildByName("TableHead")[0] is not None

        rng = self.selection.Range
        self.selection.TypeParagraph()
        self._end_range = self.selection.Range

        rows, cells = self.GetDimentions()
        self.mapper = TableMapper(self.soup, self.rows, self.columns)
        self.Table = self.selection.Tables.Add(rng,
                                               NumRows=rows,
                                               NumColumns=cells,
                                               AutoFitBehavior=constants.wdAutoFitFixed)

        for count, child in enumerate(self.GetChildren()):
            child.row_number = count + 1
            child.SetRow(self.Table.Rows(count+1))
        self.Table.Style = "Table Grid"
        self.Table.AllowAutoFit = True

    def EndRender(self):
        self.Table.Columns.AutoFit()
        self._MergeCells()
        self._end_range.Select()


class TableMapper(object):
    def __init__(self, table_parsed_html, max_rows, max_columns):
        self.table = table_parsed_html
        self.max_rows = max_rows
        self.max_columns = max_columns
        self.calculateMapping()

    def calculateMapping(self):
        mapping = OrderedDict()

        # Initialization
        for row in range(1, self.max_rows + 1):
            for column in range(1, self.max_columns + 1):
                mapping[(row, column)] = (row, column)

        rows = self.table.findAll('tr')

        # Calculating offsets
        for row_index, row in enumerate(rows, start=1):
            cells = row.findAll('td')
            for cell_index, cell in enumerate(cells, start=1):
                rowspan = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))
                #adjusts rowspan
                if rowspan != 1:
                    def is_affected_by_rowspan(position):
                        row, column = position
                        return (row_index < row < (row_index + rowspan)) and (cell_index <= column)

                    for key in filter(is_affected_by_rowspan, mapping):
                        row, column = mapping[key]
                        mapping[key] = row, column + colspan

                if colspan != 1:
                # adjusts colspan
                    def is_affected_by_colspan(position):
                        row, column = position
                        return row == row_index and cell_index < column

                    for key in filter(is_affected_by_colspan, mapping):
                        row, column = mapping[key]
                        mapping[key] = row, column + colspan - 1

        # Calculating cells to merge
        cells_to_merge = []
        for row_index, row in enumerate(rows, start=1):
            cells = row.findAll('td')
            for cell_index, cell in enumerate(cells, start=1):
                new_row_index, new_column_index = mapping[(row_index, cell_index)]

                rowspan = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))
                if (rowspan, colspan) != (1, 1):
                    cells_to_merge.append(
                        (
                            (new_row_index, new_column_index),
                            (new_row_index + rowspan - 1, new_column_index + colspan - 1)
                        )
                    )
        self.mapping = mapping
        self.cells_to_merge = cells_to_merge

class TableBody(IgnoredElement):
    pass


class TableHead(BaseElement):
    def SetRow(self, Row):
        self.GetChildren()[0].SetRow(Row)


class TableRow(BaseElement):
    AllowedChildren = ["TableCell"]
    row_number = None

    def __init__(self, *args, **kwargs):
        super(TableRow, self).__init__(*args, **kwargs)
        self.Row = None
        self.IsHeader = False

    def GetChildren(self):
        children = super(TableRow, self).GetChildren()
        map((lambda x: x.SetParent(self)), children)
        return children

    def SetRow(self, Row):
        self.Row = Row

    def StartRender(self):
        self.IsHeader = self.GetParent().GetName == "TableHead"
        mapper = self.GetParent().mapper
        for count, child in enumerate(self.GetChildren()):
            assert child.GetName() == "TableCell", "Child of TableRow is not TableCell! Its %s" % child.GetName()
            child.SetParent(self)
            new_mapping = mapper.mapping[(self.row_number, count+1)]
            child.SetCell(self.Row.Cells(new_mapping[1]))


class TableCell(BaseElement):
    StripTextAfter = True
    position = None

    def SetCell(self, Cell):
        self.Cell = Cell

    def StartRender(self):
        self.Cell.Range.Select()
    
    def HasColspan(self):
        return self._HasSpanAttribute('colspan')
    
    def GetColspan(self):
        return self._GetSpanAttribute('colspan')

    def HasRowspan(self):
        return self._HasSpanAttribute('rowspan')

    def GetRowspan(self):
        return self._GetSpanAttribute('rowspan')

    def _GetSpanAttribute(self, attribute_name):
        if attribute_name not in ('colspan', 'rowspan'):
            return None
        return int(self.GetAttrs().get(attribute_name))

    def _HasSpanAttribute(self, attribute_name):
        attr = self.GetAttrs().get(attribute_name)
        if attr:
            try:
                int_attr = int(attr)
                if int_attr > 1:
                    return True
            except ValueError:
                warnings.warn("'%s' is not a valid value for %s attribute" % (attr, attribute_name))
                return False
        return False




