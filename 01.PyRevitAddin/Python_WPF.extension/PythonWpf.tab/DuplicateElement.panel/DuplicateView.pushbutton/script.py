# -*- coding: utf-8 -*-
# ===========================
# File: script.py (Main Entry) - Duplicate Views
# ===========================
import sys
import os
import clr

# Add parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
from pyrevit import forms
from aGeneral.Logger import log_error
from Main_ViewModel import MainViewModel
from Main_Handler import ExternalEventHandler
from Autodesk.Revit.UI import ExternalEvent
from System.Windows.Data import CollectionViewSource


class IsolateWindow(forms.WPFWindow):
    def __init__(self, xaml_file, view_model):
        try:
            forms.WPFWindow.__init__(self, xaml_file)
            self.DataContext = view_model
            self.wire_events(view_model)
        except Exception as ex:
            print("Error initializing window: {}".format(str(ex)))
            log_error()
            raise

    def wire_events(self, view_model):
        """Wire up UI events"""
        try:
            # Bind DataGrid for duplicated views
            if hasattr(self, 'dgViews'):
                self.dgViews.ItemsSource = view_model.DuplicatedViews
                def _on_selection_changed(s, e):
                    try:
                        # keep VM in sync with selection
                        view_model.SelectedViewItems = list(getattr(self.dgViews, 'SelectedItems', []))
                    except:
                        pass
                self.dgViews.SelectionChanged += _on_selection_changed
                self._dg_sel_handler = _on_selection_changed

            # Find/Replace inputs
            if hasattr(self, 'tb_find'):
                self.tb_find.TextChanged += lambda s, e: setattr(view_model, 'FindText', self.tb_find.Text or "")
            if hasattr(self, 'tb_replace'):
                self.tb_replace.TextChanged += lambda s, e: setattr(view_model, 'ReplaceText', self.tb_replace.Text or "")
            # Prefix/Suffix inputs
            if hasattr(self, 'tb_prefix'):
                self.tb_prefix.TextChanged += lambda s, e: setattr(view_model, 'PrefixText', self.tb_prefix.Text or "")
            if hasattr(self, 'tb_suffix'):
                self.tb_suffix.TextChanged += lambda s, e: setattr(view_model, 'SuffixText', self.tb_suffix.Text or "")

            # Duplicate buttons
            if hasattr(self, 'bt_Duplicate'):
                self.bt_Duplicate.Click += lambda s, e: view_model.DuplicateBasicCommand.Execute(None)
            if hasattr(self, 'bt_DuplicateWithDetail'):
                self.bt_DuplicateWithDetail.Click += lambda s, e: view_model.DuplicateDetailCommand.Execute(None)
            if hasattr(self, 'bt_DuplicateAsDependent'):
                self.bt_DuplicateAsDependent.Click += lambda s, e: view_model.DuplicateDependentCommand.Execute(None)
            # New: Load selected views from Revit
            if hasattr(self, 'bt_Select'):
                self.bt_Select.Click += lambda s, e: view_model.SelectFromSelectionCommand.Execute(None)

            # Rename button
            if hasattr(self, 'bt_Rename'):
                self.bt_Rename.Click += lambda s, e: view_model.RenameCommand.Execute(None)
            # New: Remove selected rows
            if hasattr(self, 'bt_Remove'):
                self.bt_Remove.Click += lambda s, e: view_model.RemoveCommand.Execute(None)

            # Close button
            if hasattr(self, 'bt_Cancel'):
                self.bt_Cancel.Click += lambda s, e: self.on_close_clicked(s, e)

            # store VM
            self._view_model = view_model

        except Exception as ex:
            print("Error wiring events: {}".format(str(ex)))

    def on_close_clicked(self, sender, e):
        try:
            self.Close()
        except Exception as ex:
            print("Error closing window: {}".format(str(ex)))

def show_window():
    try:
        handler = ExternalEventHandler()
        external_event = ExternalEvent.Create(handler)
        view_model = MainViewModel(external_event, handler)
        handler.view_model = view_model

        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, "UI-DuplicateView.xaml")
        window = IsolateWindow(xaml_path, view_model)
        window.Show()
    except Exception as ex:
        print("Error showing window: {}".format(str(ex)))
        log_error()

if __name__ == '__main__':
    show_window()