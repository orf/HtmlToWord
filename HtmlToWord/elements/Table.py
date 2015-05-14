from collections import OrderedDict

from HtmlToWord.elements.Base import *
from win32com.client import constants


class Table(BlockElement):
    """
    I seem to support merged cells and rows now, but in their infinite complexity I am
    sure someone will find ways to break me. I support two types of table:
     * Simple ones with no tbody or thead (just <table><tr><td>...)
     * Complex ones with a tbody and thead element.
    """
    AllowedChildren = ["TableRow", "TableBody", "TableHead"]

    def __init__(self, *args, **kwargs):
        super(Table, self).__init__(*args, **kwargs)
        self.TableRow = 0
        self.HasHeader = False

    def _MergeCells(self):
        """
        Merge cells in order to deal with 'colspan' and 'rowspan' attributes.
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
         This needs to be done here, after table has been rendered (otherwise the merged
         cell will have the size of the original cell), and in reversed order (from bottom/right to top/left)
         to work (after merging, the word Cell object's RowIndex and ColumnIndex properties values may change)
         """
        for originCellCoordinates, targetCellCoordinates in reversed(self.mapper.cells_to_merge):
            originCell = self.Table.Cell(*originCellCoordinates)
            targetCell = self.Table.Cell(*targetCellCoordinates)
            originCell.Merge(targetCell)

    def StartRender(self):
        self.HasHeader = self.GetChildByName("TableHead")[0] is not None

        rng = self.selection.Range
        self.selection.TypeParagraph()
        self._end_range = self.selection.Range

        self.mapper = TableMapper(self.soup)
        rows, cells = self.mapper.max_rows, self.mapper.max_columns
        self.Table = self.selection.Tables.Add(rng,
                                               NumRows=rows,
                                               NumColumns=cells,
                                               AutoFitBehavior=constants.wdAutoFitWindow)

        for count, child in enumerate(self.GetChildren()):
            row_number = count + 1
            child.SetRow(self.Table.Rows(row_number), row_number)
        self.Table.Style = "Table Grid"

    def ApplyFormatting(self, start_pos, end_pos):
        super(Table, self).ApplyFormatting(start_pos, end_pos)
        if 'border' in self.attrs:
            border = self.attrs['border']
            if border == '0':
                self.Table.Borders.Enable = 0

    def EndRender(self):
        self._MergeCells()
        self._end_range.Select()


class TableMapper(object):
    def __init__(self, table_parsed_html):
        self.calculateMapping(table_parsed_html)

    def calculateMapping(self, table_parsed_html):
        mapping = OrderedDict()
        tablerows = table_parsed_html.findAll('tr')

        # Calculating offsets
        for row_index, row in reversed(list(enumerate(tablerows, start=1))):
            cells = row.findAll(('td', 'th'))
            for cell_index, cell in reversed(list(enumerate(cells, start=1))):
                mapping[(row_index, cell_index)] = (row_index, cell_index)

                rowspan = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))

                if rowspan != 1:  # Adjust rowspan
                    def is_affected_by_rowspan(position):
                        row, column = position
                        return (row_index < row < (row_index + rowspan)) and (cell_index <= column)

                    for key in filter(is_affected_by_rowspan, mapping):
                        row, column = mapping[key]
                        mapping[key] = row, column + colspan

                if colspan != 1:  # Adjust colspan
                    def is_affected_by_colspan(position):
                        row, column = position
                        return row == row_index and cell_index < column

                    for key in filter(is_affected_by_colspan, mapping):
                        row, column = mapping[key]
                        mapping[key] = row, column + colspan - 1

        # Calculate the maximum table size
        rows, columns = zip(*mapping.values())
        self.max_rows, self.max_columns = max(rows), max(columns)

        # Calculating cells to merge
        cells_to_merge = []
        for row_index, row in enumerate(tablerows, start=1):
            cells = row.findAll(('td', 'th'))
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


class TableHead(IgnoredElement):
    pass


class TableRow(BaseElement):
    AllowedChildren = ["TableCell"]

    def getStartPosition(self):
        return self.Row.Range.Start

    def getEndPosition(self):
        return self.Row.Range.End

    def SetRow(self, Row, row_number):
        self.Row = Row
        self.row_number = row_number

    def StartRender(self):
        parent = self.GetParent()
        mapping = parent.mapper.mapping
        for count, child in enumerate(self.GetChildren()):
            assert child.GetName() == "TableCell", "Child of TableRow is not TableCell! Its %s" % child.GetName()
            new_row, new_column = mapping[(self.row_number, count+1)]
            child.SetCell(self.Row.Cells(new_column))


class TableCell(BaseElement):
    StripTextAfter = True
    position = None

    def getStartPosition(self):
        return self.Cell.Range.Start

    def getEndPosition(self):
        return self.Cell.Range.End

    def SetCell(self, Cell):
        self.Cell = Cell

    def StartRender(self):
        self.Cell.Range.Select()

