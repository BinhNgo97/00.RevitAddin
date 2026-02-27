# -*- coding: utf-8 -*-
class ExcelViewModel(object):

    def __init__(self):
        self.selected_elements = []       # danh sách ElementId
        self.unique_parameters = []       # danh sách parameter duy nhất
        self.data_table = []              # dữ liệu cho DataGrid
        self.excel_path = None
