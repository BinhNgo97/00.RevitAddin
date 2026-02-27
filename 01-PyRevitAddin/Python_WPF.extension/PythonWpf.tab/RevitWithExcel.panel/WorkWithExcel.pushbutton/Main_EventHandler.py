# -*- coding: utf-8 -*-
import clr
clr.AddReference("Microsoft.Win32")
from Microsoft.Win32 import OpenFileDialog

class ExcelEventHandler(object):

    def __init__(self, ui, vm, handler, uidoc):
        self.ui = ui
        self.vm = vm
        self.h = handler
        self.uidoc = uidoc

        # Gán event button
        self.ui.btnSelect.Click += self.on_select
        self.ui.btnFile.Click += self.on_select_file
        self.ui.btnExport.Click += self.on_export
        self.ui.btnRefresh.Click += self.on_refresh


    # ---------------- BUTTON: SELECT IN REVIT ----------------
    def on_select(self, sender, e):
        self.h.select_elements(self.uidoc)

        self.ui.lvParameters.Items.Clear()
        for p in self.vm.unique_parameters:
            self.ui.lvParameters.Items.Add(p)


    # ---------------- BUTTON: SELECT EXCEL FILE ----------------
    def on_select_file(self, sender, e):
        dlg = OpenFileDialog()
        dlg.Filter = "Excel files (*.xlsx)|*.xlsx"
        if dlg.ShowDialog():
            self.h.set_excel_path(dlg.FileName)


    # ---------------- BUTTON: EXPORT DATA ----------------
    def on_export(self, sender, e):
        selected_params = [item for item in self.ui.lvParameters.SelectedItems]
        self.h.export_to_excel(selected_params)


    # ---------------- BUTTON: REFRESH DATA ----------------
    def on_refresh(self, sender, e):
        self.h.refresh_from_excel()

        self.ui.dgData.Columns.Clear()
        self.ui.dgData.Items.Clear()

        if not self.vm.data_table:
            return

        # tạo cột
        for key in self.vm.data_table[0].keys():
            col = System.Windows.Controls.DataGridTextColumn()
            col.Header = key
            col.Binding = System.Windows.Data.Binding(key)
            self.ui.dgData.Columns.Add(col)

        # thêm dòng
        for row in self.vm.data_table:
            self.ui.dgData.Items.Add(row)
