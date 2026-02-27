# -*- coding: utf-8 -*-
# Entry for "Copy Family & Type from RVT"

import os, time, clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')

from pyrevit import forms
from Autodesk.Revit.UI import ExternalEvent
from Main_EventHandler import ExternalEventHandler

class CategoryRow(object):
    def __init__(self, name, bic):
        self.Name = name
        self.BIC = int(bic)

class TypeRow(object):
    def __init__(self, name, uid):
        self.Name = name
        self.UniqueId = uid

class CopyTypesWindow(forms.WPFWindow):
    def __init__(self, xaml):
        forms.WPFWindow.__init__(self, xaml)
        self._handler = ExternalEventHandler()
        self._ext = ExternalEvent.Create(self._handler)
        self._source_file = None
        # wire events
        self.btnFilepath.Click += self.on_pick_file
        self.cmbCategory.SelectionChanged += self.on_category_changed
        self.btnTransfer.Click += self.on_transfer

    def _await(self):
        while self._handler.is_busy:
            time.sleep(0.05)

    def on_pick_file(self, s, e):
        path = forms.pick_file(file_ext="rvt", restore_dir=True, title="Select source Revit file")
        if not path:
            return
        self._source_file = path
        self.btnFilepath.Content = path
        self._handler.source_file = path
        self._handler.action = "load_categories"
        self._handler.is_busy = True           # avoid race
        self._ext.Raise()
        self._await()
        cats = self._handler.categories or []
        rows = [CategoryRow(x.get('Name'), x.get('BIC')) for x in cats]
        self.cmbCategory.ItemsSource = rows
        self.cmbCategory.DisplayMemberPath = "Name"
        if rows:
            self.cmbCategory.SelectedIndex = 0
        self.lstTypes.ItemsSource = None

    def on_category_changed(self, s, e):
        item = self.cmbCategory.SelectedItem
        if not item or not self._source_file:
            return
        self._handler.source_file = self._source_file
        self._handler.category_bic = int(getattr(item, 'BIC', 0))
        self._handler.action = "load_types"
        self._handler.is_busy = True           # avoid race
        self._ext.Raise()
        self._await()
        rows = [TypeRow(x['Name'], x['UniqueId']) for x in (self._handler.types or [])]
        self.lstTypes.ItemsSource = rows

    def on_transfer(self, s, e):
        sel = list(self.lstTypes.SelectedItems) if self.lstTypes.SelectedItems else []
        if not sel:
            forms.alert("No types selected.", title="Copy Types")
            return
        self._handler.source_file = self._source_file
        self._handler.selected_unique_ids = [x.UniqueId for x in sel]
        self._handler.action = "transfer"
        self._handler.is_busy = True           # avoid race
        self._ext.Raise()
        self._await()
        forms.alert(self._handler.message or "Done.", title="Copy Types")

def main():
    xaml_path = os.path.join(os.path.dirname(__file__), "UI.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("UI.xaml not found:\n{}".format(xaml_path))
        return
    CopyTypesWindow(xaml_path).ShowDialog()

if __name__ == "__main__":
    main()