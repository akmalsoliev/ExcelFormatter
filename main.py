import openpyxl
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter

import numpy as np 


class Excel_Fomatter:
    def __init__(self, file_path:str, sheet:str, index_size:int=1, column_size:int=1):
        """
        Produces formated Excel file.
        :param file_path: Full path of the excel file.
        :param sheet: Sheet that will be formatted.
        :param index_size: if working with multilayer index then specify the size of it (how many columns is the multilayer index).
        :param column_size: if working with multilayer column then specify the size of it (how many rows is the multilayer column).
        """
        self.file_path = file_path
        self.sheet = sheet
        self.index_size = index_size
        self.column_size = column_size
        self.workbook = self.get_workbook()
        self.worksheet = self.get_worksheet()
        self.xl_cols, self.xl_rows = self.get_shape()

    def get_workbook(self):
        return openpyxl.load_workbook(self.file_path)

    def get_worksheet(self):
        return self.workbook[self.sheet]

    def get_shape(self):
        xl_rows, xl_cols = np.array([*self.worksheet.columns]).shape
        return xl_rows, xl_cols

    def del_row(self, row, column):
        self.worksheet.delete_rows(row, column)
        return "Row deletion complete!"

    def set_column_percentage(self, start_col:int, start_row:int, steps:int, skip_rows:int=1):
        """
        Formats columns to percentage
        :param start_col: Starting column.
        :param steps: Every X column will be formatted as percentage.
        :param skip_rows: Number of rows that should be skipped in the formatting (number of header rows).
        """

        for column in np.arange(start_col, self.xl_cols+1, steps):
            for row in np.arange(start_row, self.xl_rows):
                if row > skip_rows:
                    cell = self.worksheet.cell(row=row, column=column)
                    setattr(cell, "style", "Percent")
        return "Percentage formatting is complete!"

    def freeze_panes(self, cell:str=None, row:int=None, column:int=None):
        if cell and (row or column):
            raise """Need to specify either `cell` or `row` and `column`. Cannot declare `cell` and `row` or `column`."""
        elif cell:
            self.worksheet.freeze_panes = cell
            return "Pane freeze is complete"
        elif row and column:
            ws = self.worksheet
            selected_cell = ws.cell(row, column)
            self.worksheet.freeze_panes = selected_cell
            return "Pane freeze is complete"
        else:
            raise "Need to specify either `cell` or `row` and `column`"

    def fill_color(self, color:str):
            return PatternFill(start_color=color, end_color=color, fill_type='solid')

    def color_columns(self, start:int, steps:int, color0:str="FFFFFF", color1:str="EEEBFF"):
        # Color Fill
        col_fill0 = self.fill_color(color0)
        col_fill1 = self.fill_color(color1)

        start_col_range = np.arange(start, self.xl_cols, steps)
        end_col_range = np.arange(start+steps, self.xl_cols+steps, steps)
        cols = [x for x in self.worksheet.columns]

        for i, (start, end) in enumerate(zip(start_col_range, end_col_range)):
            for cell in np.array(cols[start:end]).flatten():
                if i%2:
                    cell.fill = col_fill0
                else:
                    cell.fill = col_fill1
        return "Column coloring is complete"

    def header_color(self, color:str="B4C6E7", header_num_rows:int=2):
        head_fill = self.fill_color(color)
        try:
            for row in self.worksheet["1:{}".format(header_num_rows)]:
                for cell in row:
                    cell.fill = head_fill
        except TypeError:
            for cell in self.worksheet[header_num_rows]:
                cell.fill = head_fill
        return "Header coloring is complete"

    def color_row(self, start_row:int, end_row:int, color:str="FCD5B4"):
        color_fill = self.fill_color(color)
        if start_row - end_row == 0:
            for cell in self.worksheet["{}:{}".format(start_row, end_row)]:
                cell.fill = color_fill
        else:
            for row in self.worksheet["{}:{}".format(start_row, end_row)]:
                for cell in row:
                    cell.fill = color_fill
        return "Row coloring is complete"

    def format_border(self):
        border = Border(left=Side(border_style='thin', color='000000'),
                        right=Side(border_style='thin', color='000000'),
                        top=Side(border_style='thin', color='000000'),
                        bottom=Side(border_style='thin', color='000000'))
        [setattr(cell, "border", border) for cell in np.array([*self.worksheet.columns]).flatten()]
        return "Border formatting is complete"

    def columns_width(self, padding:int=3):

        dim_holder = DimensionHolder(worksheet=self.worksheet)

        for col in np.arange(1, self.xl_cols + 1):
            # intiating the width of the column as 0, as it will serve base size.
            width = 0
            for row in np.arange(1, self.xl_rows + 1):
                cell_value = self.worksheet.cell(column=col, row=row).value
                if cell_value:
                    # finding what's the len of the cell's values and then setting it
                    # as a new base `width`
                    cell_len = len(str(cell_value))
                    if cell_len > width:
                        width = cell_len + padding

            dim_holder[get_column_letter(col)] = (
                ColumnDimension(
                    self.worksheet,
                    min=col,
                    max=col,
                    width=width)
            )

        self.worksheet.column_dimensions = dim_holder
        return "Column width adjustment is complete"

    def save(self):
        self.workbook.save(self.file_path)
        return "File saving is complete"
