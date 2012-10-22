from xlwt import Workbook,easyxf,Formula
from datetime import date
import sys


# class to save sheet name, headers and data format in each sheet
class MySheet:
    def __init__(self, sheet_name):
        self._name = sheet_name
        self._headers = {}
        self._data_format = {}
        self._is_row_empty = []


# main class
class ExcelFile:
    def __init__(self, file_path, header_format, data_format):
        self._file_path = file_path
        self._header_style0 = easyxf('borders: left thick;' + header_format)
        self._header_style1 = easyxf(header_format)
        self._data_style0 = easyxf('borders: left thick;' + data_format)
        self._data_style1 = easyxf(data_format)
        self._date_style0 = easyxf('borders: right thin, left thick; alignment: horizontal left, wrap true;', num_format_str='DD-MM-YYYY')
        self._date_style1 = easyxf('borders: right thin; alignment: horizontal left, wrap true;', num_format_str='DD-MM-YYYY')
        self._empty_cell_style0 = easyxf('borders: right thin, left thick;')
        self._empty_cell_style1 = easyxf('borders: right thin;')
        self._last_row_style0 = easyxf('borders: right thin, bottom thin, left thick')
        self._last_row_style1 =  easyxf('borders: right thin, bottom thin')
        self._workbook = Workbook()
        self._sheets = [] # contains the sheet objects
        self._current_sheet = None # current sheet object of class 'Workbook'
        self._my_sheet = None # current sheet object of class 'MySheet'

    # add a new sheet in the workbook    
    def add_sheet(self, sheet_name):
        self._current_sheet = self._workbook.add_sheet(sheet_name, cell_overwrite_ok = True)
        self._current_sheet.show_grid = 0 # removes all the grid lines in excel
        self._my_sheet = MySheet(sheet_name)
        self._sheets.append(self._my_sheet)

    # set a sheet as current sheet
    def set_sheet(self, sheet_name):
        index = 0 # integer value representing each sheet
        flag = False
        for sheet in self._sheets:
            if sheet._name == sheet_name:
                flag = True
                self._my_sheet = sheet
                break
            index += 1
        if flag == True:
            self._current_sheet = self._workbook.get_sheet(index)
        elif flag == False:
            raise ValueError('Invalid sheet name.')

    # set a new header format into all sheets
    def set_header_format(self, header_format):
        self._header_style0 = easyxf('borders: left thick;' + header_format)
        self._header_style1 = easyxf(header_format)

    # add a new format into a column of current sheet
    def set_data_format(self, header_name, column_format):
        column_no = self._check_column_is_name(header_name)
        if column_no != None:
            if column_format == 'date':
                self._my_sheet._data_format[column_no] = column_format
            elif column_no == 0:
                self._my_sheet._data_format[column_no] = easyxf('borders: right thin, left thick; alignment: horizontal left, wrap true;' + column_format)
            else:
                self._my_sheet._data_format[column_no] = easyxf('borders: right thin; alignment: horizontal left, wrap true;' + column_format)
        else:
            raise ValueError('Invalid header name')

    # write a single header into current sheet
    def write_header(self, column_no, header_name):
        if column_no == 0:
            style = self._header_style0
        else:
            style = self._header_style1
        self._current_sheet.write(0, column_no, header_name, style)
        self._my_sheet._headers[column_no] = header_name

    # write a list of headers into current sheet
    def write_headers(self, headers):
        for header in headers:
            self.write_header(headers.index(header), header)

    # write data into a single cell of current sheet
    def write_data(self, row_no, column, data, column_is_name = None):
        # If 'column_is_name' argument is True then 'column' argument contains header-name
        # If 'column_is_name' argument is False then 'column' argument contains column-number
        column_no = None
        if column_is_name == True:
            column_no = self._check_column_is_name(column)
        elif column_is_name == False:
            column_no = self._check_column_is_number(column)
        elif isinstance(column, int):
            column_no = self._check_column_is_name(column)
            if column_no != None:
                raise ConflictError(
                    """The 'column' argument in function write_data() creates a conflict. 
                    Set 'column_is_name' argument as 'True' if 'column' argument is the header-name or 
                    set 'column_is_name' argument as 'False' if 'column' argument is the column-number"""
                    ) 
            else:
                column_no = self._check_column_is_number(column)
        else:
            column_no = self._check_column_is_name(column)

        if column_no == None:
            raise ValueError('Invalid header name or column number.')

        # set style in empty cells
        if row_no not in self._my_sheet._is_row_empty:
            self._set_empty_cell_style(row_no)
            self._my_sheet._is_row_empty.append(row_no)

        # select the appropriate style and write data
        column_format = self._my_sheet._data_format.get(column_no)
        style = self._set_data_style(column_no, column_format)
        self._current_sheet.write(row_no, column_no, data, style)

    # set style for data cells
    def _set_data_style(self, column_no, column_format):
        if column_format == None:
            if column_no == 0:
                style = self._data_style0
            else:
                style = self._data_style1
        elif column_format == 'date':
            self._current_sheet.col(column_no).width = 256 * 12 # 12 spaces for date cell
            if column_no == 0:
                style = self._date_style0
            else:
                style = self._date_style1
        else:
            style = column_format
        return style

    # set style into empty cells of a row
    def _set_empty_cell_style(self, row_no):
        for column_no in range(len(self._my_sheet._headers)):
            if column_no == 0:
                self._current_sheet.write(row_no, column_no, None, self._empty_cell_style0)
            else:
                self._current_sheet.write(row_no, column_no, None, self._empty_cell_style1)

    # check if the column given is header name
    def _check_column_is_name(self, column):
        column_no = None
        if column in self._my_sheet._headers.values():
            column_no = [key for key, value in self._my_sheet._headers.iteritems() if value == column][0]
        return column_no

    # check if the column given is column number
    def _check_column_is_number(self, column):
        column_no = None
        if column in self._my_sheet._headers:
            column_no = column
        return column_no

    # write data into next available row of current sheet
    def write_row(self, row_data):

        if len(self._my_sheet._headers) < len(row_data):
            raise ExcessDataError("The argument 'row_data' in function write_row() contains more data than the number of headers")

        # set style in empty cells
        row_no = self._current_sheet.last_used_row + 1
        if row_no not in self._my_sheet._is_row_empty:
            self._set_empty_cell_style(row_no)
            self._my_sheet._is_row_empty.append(row_no)

        row = self._current_sheet.row(row_no)
        for value in row_data:
            column_no = row_data.index(value) 
            column_format = self._my_sheet._data_format.get(column_no)
            style = self._set_data_style(column_no, column_format)
            row.write(column_no, value, style)

    # call this function at the end to save the excel file
    def save(self):
        self._set_last_row_style()
        self._workbook.save(self._file_path)

    # write last row into all sheets to complete the styling
    def _set_last_row_style(self):
        index = 0
        for sheet in self._sheets:
            current_sheet = self._workbook.get_sheet(index)
            index += 1
            row = current_sheet.row(current_sheet.last_used_row + 1)
            for i in range(len(sheet._headers)):
                if i == 0:
                    row.write(i, None, self._last_row_style0)
                else:
                    row.write(i, None, self._last_row_style1)


class ConflictError(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr(self.value)


class ExcessDataError(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr(self.value)


if __name__ == '__main__':

    excel_file = ExcelFile(
                           'excel_file.xls', 
                           '''font: name Arial, bold True; 
                           alignment: horizontal center, wrap true;
                           borders: top thick, right thin, bottom thin; 
                           pattern: pattern solid, fore_colour light_green;''',
                           '''font: name Arial; 
                           alignment: horizontal left, wrap true;
                           borders: right thin'''
                          )

    today = date.today()
    link = 'HYPERLINK("temp/hello.txt";"Sample File")'

    excel_file.add_sheet('sheet1')
    excel_file.add_sheet('sheet2')

    excel_file.set_sheet('sheet1')
    excel_file.write_headers(['Name', 'Age', 'Date', 'File'])
    excel_file.set_data_format('Date', 'date')
    excel_file.set_data_format('Age', 'font: name Arial, bold True, color red')
    excel_file.write_data(1, 0, 'John')
    excel_file.write_data(1, 'Age', 22)
    excel_file.write_data(2, 0, 'Forest')
    excel_file.write_row(['Ben', 34.122, today, Formula(link)])
    excel_file.write_row(['Liz'])
    excel_file.write_row(['Jennifer', 0000])
    excel_file._current_sheet.flush_row_data()

    excel_file.set_sheet('sheet2')
    excel_file.write_headers(['Name', 'Age', 'Date'])
    excel_file.set_data_format('Name', 'font: name Arial, bold True, color red')
    excel_file.write_data(1, 'Name', 'Karen')
    excel_file.write_data(1, 1, 50)
    excel_file.write_data(1, 'Date', today)
    excel_file.write_data(2, 'Name', 'Shawn')
    excel_file.write_data(3, 'Name', 'Loran')
    excel_file.write_row(['Bane', 45])
    excel_file._current_sheet.flush_row_data()

    excel_file.save()