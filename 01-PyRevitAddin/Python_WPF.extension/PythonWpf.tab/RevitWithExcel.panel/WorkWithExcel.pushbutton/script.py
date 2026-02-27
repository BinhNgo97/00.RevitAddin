# -*- coding: utf-8 -*-
# ===========================
# File: script.py (Main Entry) - Duplicate Element Type
# ===========================
import sys
import os
import clr
import time

# Add parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
from pyrevit import forms
from pyrevit import revit
from aGeneral.Logger import log_error
from Main_ViewModel import MainViewModel
# from Main_Handler import ExternalEventHandler
from Main_EventHandler import ExternalEventHandler
from Autodesk.Revit.UI import ExternalEvent
import System.Guid

# Keep a global reference so the modeless window isn't GC'd
_modeless_window = None

class DuplicateElementWindow(forms.WPFWindow):
    def __init__(self, xaml_file, view_model):
        try:
            forms.WPFWindow.__init__(self, xaml_file)
            self.DataContext = view_model
            self._view_model = view_model
            try:
                self._view_model.PropertyChanged += self.on_vm_property_changed
            except:
                pass
            self.wire_events(view_model)
        except Exception as ex:
            # print("Error initializing window: {}".format(str(ex)))
            log_error()
            raise

    def wire_events(self, view_model):
        """Wire up UI events"""
        try:
            # ComboBox selection events
            if hasattr(self, 'cmbCategories'):
                if callable(getattr(self, 'on_category_selection_changed', None)):
                    self.cmbCategories.SelectionChanged += self.on_category_selection_changed
                self.cmbCategories.ItemsSource = view_model.Categories
                self.cmbCategories.DisplayMemberPath = "Name"
                try:
                    if view_model.Categories.Count > 0:
                        self.cmbCategories.SelectedIndex = 0
                except:
                    pass

            # unified Fa/El dropdown
            if hasattr(self, 'cmbFaElType'):
                if callable(getattr(self, 'on_fael_selection_changed', None)):
                    self.cmbFaElType.SelectionChanged += self.on_fael_selection_changed
                self.cmbFaElType.ItemsSource = view_model.FaElTypes
                self.cmbFaElType.DisplayMemberPath = "Name"

            if hasattr(self, 'cmbSheets'):
                if callable(getattr(self, 'on_sheet_selection_changed', None)):
                    self.cmbSheets.SelectionChanged += self.on_sheet_selection_changed
                self.cmbSheets.ItemsSource = view_model.SheetNames

            # ImportData button (on-demand import to grid)
            if hasattr(self, 'btnImportData'):
                self.btnImportData.Click += lambda s, e: view_model.ImportDataCommand.Execute(None)
            # existing refresh still available (re-imports current sheet)
            if hasattr(self, 'btnRefresh'):
                self.btnRefresh.Click += lambda s, e: view_model.RefreshDataCommand.Execute(None)
            # Wire bulk actions
            if hasattr(self, 'btnUpdateSelected'):
                self.btnUpdateSelected.Click += lambda s, e: view_model.UpdateAllCommand.Execute(None)
            if hasattr(self, 'btnCreateAll'):
                self.btnCreateAll.Click += lambda s, e: view_model.CreateAllNewCommand.Execute(None)
            # Browse Excel file button
            if hasattr(self, 'btnSelectExcel'):
                self.btnSelectExcel.Click += lambda s, e: view_model.SelectExcelCommand.Execute(None)
                # initialize button content with current path or placeholder
                try:
                    self.btnSelectExcel.Content = view_model.ExcelFilePath or r"C:\Users..."
                except:
                    pass

            if hasattr(self, 'btnCreateType'):
                self.btnCreateType.Click += lambda s, e: view_model.CreateTypeCommand.Execute(None)
                
            if hasattr(self, 'btnClose'):
                if callable(getattr(self, 'on_close_clicked', None)):
                    self.btnClose.Click += self.on_close_clicked
                
            # DataGrid auto column handling: hide Status & Action (we have a template Action column)
            if hasattr(self, 'dataGrid'):
                try:
                    self.dataGrid.AutoGeneratingColumn += self.on_datagrid_autogeneratingcolumn
                except:
                    pass
            # Text bindings
        except Exception as ex:
            # print("Error wiring events: {}".format(str(ex)))
            pass

    def on_datagrid_autogeneratingcolumn(self, sender, e):
        """Cancel auto-generated columns we don't want to display."""
        try:
            header = ""
            try:
                header = str(e.Column.Header) if hasattr(e, "Column") and hasattr(e.Column, "Header") else ""
            except:
                header = ""
            # Hide Status (computed) and Action (we have a template column)
            if header in ("Status", "Action"):
                e.Cancel = True
                return
            # Ensure validation is enabled so DataRow column errors are picked up per cell
            try:
                from System.Windows.Controls import DataGridBoundColumn
                if isinstance(e.Column, DataGridBoundColumn):
                    b = e.Column.Binding
                    if hasattr(b, "ValidatesOnDataErrors"):
                        b.ValidatesOnDataErrors = True
                    if hasattr(b, "NotifyOnValidationError"):
                        b.NotifyOnValidationError = True
            except:
                pass
        except Exception as ex:
            # print("AutoGeneratingColumn handler error: {}".format(str(ex)))
            pass

    def on_type_selection_changed(self, sender, e):
        """Handle type selection change"""
        try:
            if sender.SelectedItem:
                self._view_model.SelectedType = sender.SelectedItem
        except Exception as ex:
            # print("Error in type selection: {}".format(str(ex)))
            pass

    def on_category_selection_changed(self, sender, e):
        try:
            if sender.SelectedItem:
                self._view_model.SelectedCategory = sender.SelectedItem
                # Sau khi VM load Families, nếu có Fa-Type thì chọn mục đầu để load El-Type
                try:
                    if hasattr(self, 'cmbFamilies') and self.cmbFamilies.Items.Count > 0:
                        self.cmbFamilies.SelectedIndex = 0
                except: 
                    pass
        except Exception as ex:
            # print("Error in category selection: {}".format(str(ex)))
            pass

    def on_family_selection_changed(self, sender, e):
        try:
            if sender.SelectedItem:
                self._view_model.SelectedFamily = sender.SelectedItem
                # Sau khi VM load ElementTypes, auto chọn El-Type đầu nếu có
                try:
                    if hasattr(self, 'cmbTypes') and self.cmbTypes.Items.Count > 0:
                        self.cmbTypes.SelectedIndex = 0
                except Exception as ex:
                    # print("Error auto-selecting element type: {}".format(str(ex)))
                    pass
        except Exception as ex:
            # print("Error in family selection: {}".format(str(ex)))
            pass

    def on_fael_selection_changed(self, sender, e):
        """Handle selection change of unified Fa/El dropdown"""
        try:
            if sender.SelectedItem:
                # set SelectedFaEl on VM
                try:
                    self._view_model.SelectedFaEl = sender.SelectedItem
                except Exception as ex:
                    # print("Error setting SelectedFaEl: {}".format(str(ex)))
                    pass
        except Exception as ex:
            # print("Error in fa/el selection: {}".format(str(ex)))
            pass

    def on_sheet_selection_changed(self, sender, e):
        """Handle sheet selection change"""
        try:
            if sender.SelectedItem:
                self._view_model.SelectedSheet = sender.SelectedItem
        except Exception as ex:
            # print("Error in sheet selection: {}".format(str(ex)))
            pass

    def on_vm_property_changed(self, sender, e):
        """Reflect simple VM property updates on UI elements that aren't XAML-bound."""
        try:
            pname = getattr(e, 'PropertyName', '')
            if pname == "ExcelFilePath" and hasattr(self, 'btnSelectExcel'):
                self.btnSelectExcel.Content = self._view_model.ExcelFilePath or r"C:\Users..."
            # Update bottom counters whenever the grid's DataTable changes (Refresh/Import)
            if pname == "DataTableDefaultView":
                try:
                    existing, update, new = self._view_model.get_status_counts()
                    if hasattr(self, 'btnExisting'):
                        self.btnExisting.Content = "Existing: {}".format(existing)
                    if hasattr(self, 'btnUpdate'):
                        self.btnUpdate.Content = "Need Update: {}".format(update)
                    if hasattr(self, 'btnNew'):
                        self.btnNew.Content = "New: {}".format(new)
                except:
                    pass
        except:
            pass

def show_window():
    """Show modeless window so user can keep working in Revit."""
    global _modeless_window
    try:
        # Create external event handler
        handler = ExternalEventHandler()
        external_event = ExternalEvent.Create(handler)

        # Create view model
        view_model = MainViewModel(external_event, handler)
        handler.view_model = view_model

        # XAML path
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, "UI-DupE.xaml")
        if not os.path.exists(xaml_path):
            try:
                # forms.alert("UI-DupE.xaml not found:\n{}".format(xaml_path), title="Error")  # disabled popup
                pass
            except:
                print("UI-DupE.xaml not found: {}".format(xaml_path))
            return

        # Create window and show modeless
        win = DuplicateElementWindow(xaml_path, view_model)
        # Clear global when closed
        try:
            def _on_closed(sender, e):
                global _modeless_window
                _modeless_window = None
            win.Closed += _on_closed
        except:
            pass

        win.Show()
        _modeless_window = win
        # print("Modeless window shown")
    except Exception as ex:
        # print("Error showing window: {}".format(str(ex)))
        log_error()
        try:
            # forms.alert("Error showing Duplicate Element window:\n{}".format(str(ex)), title="Error")  # disabled popup
            pass
        except:
            pass

def main():
    # Use modeless show
    show_window()

# Ensure main() runs when executed by pyRevit as well as direct run
if __name__ == '__main__':
    main()
else:
    try:
        # If script run inside pyRevit runner, call main() to show UI
        if 'pyrevit' in sys.modules:
            main()
    except Exception:
        pass