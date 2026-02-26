# -*- coding: utf-8 -*-
# ===========================
# File: script.py (Main Entry) - Isolate Element By Parameter
# ===========================
import sys
import os
import clr

# Add parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
# NEW: access Revit UI events
clr.AddReference('RevitAPIUI')
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
            # ListView with checkboxes: just bind ItemsSource; toggling handled by ViewModel
            if hasattr(self, 'lvSearchParameters'):
                self.lvSearchParameters.ItemsSource = view_model.ParameterNames
                # Apply filter view bound to FilterText
                try:
                    self._param_view = CollectionViewSource.GetDefaultView(view_model.ParameterNames)
                    def _filter(item):
                        try:
                            text = (view_model.FilterText or "").lower()
                            if not text:
                                return True
                            return text in (item.Name or "").lower()
                        except:
                            return True
                    self._filter_func = _filter
                    self._param_view.Filter = self._filter_func

                    # React to FilterText changes
                    def _on_vm_prop_changed(sender, args):
                        try:
                            if args.PropertyName == 'FilterText':
                                self._param_view.Refresh()
                        except:
                            pass
                    view_model.PropertyChanged += _on_vm_prop_changed
                    self._vm_prop_handler = _on_vm_prop_changed
                except Exception as ex:
                    print('Error setting up filter: {}'.format(str(ex)))

            # DataGrid: parameter rows
            if hasattr(self, 'dgParameterValues'):
                self.dgParameterValues.ItemsSource = view_model.ParameterRows
                # keep selection in VM
                def _on_selection_changed(s, e):
                    try:
                        view_model.SelectedParameterRow = s.SelectedItem
                    except:
                        pass
                self.dgParameterValues.SelectionChanged += _on_selection_changed
                self._dg_sel_handler = _on_selection_changed
                # double-click to remove selected row
                def _on_double_click(s, e):
                    try:
                        if self.dgParameterValues.SelectedItem:
                            view_model.remove_selected_row(None)
                    except:
                        pass
                self.dgParameterValues.MouseDoubleClick += _on_double_click
                self._dg_double_click = _on_double_click

            # Buttons
            if hasattr(self, 'btnIsolate'):
                self.btnIsolate.Click += lambda s, e: view_model.IsolateCommand.Execute(None)
            if hasattr(self, 'btnUnIsolate'):
                self.btnUnIsolate.Click += lambda s, e: view_model.UnIsolateCommand.Execute(None)
            if hasattr(self, 'btnCopyRow'):
                self.btnCopyRow.Click += lambda s, e: view_model.DuplicateRowCommand.Execute(None)

            # Store view model reference
            self._view_model = view_model

            # NEW: refresh values when active view changes
            def _on_view_activated(sender, e):
                try:
                    vm = getattr(self, '_view_model', None)
                    if not vm:
                        return
                    # get current doc/view safely
                    try:
                        from pyrevit import revit as _rv
                        doc = _rv.doc
                        view = _rv.active_view
                    except:
                        # fallback
                        doc = __revit__.ActiveUIDocument.Document if __revit__.ActiveUIDocument else None
                        view = __revit__.ActiveUIDocument.ActiveView if __revit__.ActiveUIDocument else None
                    vm.refresh_after_view_change(doc, view, None)
                except:
                    pass

            # subscribe and keep handler ref for unsubscribe
            __revit__.ViewActivated += _on_view_activated
            self._view_activated_handler = _on_view_activated
        except Exception as ex:
            print("Error wiring events: {}".format(str(ex)))

    # Checkbox toggling is two-way bound; no handler needed

    # no extra handlers needed

    def on_close_clicked(self, sender, e):
        """Handle close button click"""
        try:
            # NEW: unsubscribe from ViewActivated to avoid leaks
            try:
                if hasattr(self, '_view_activated_handler') and self._view_activated_handler:
                    __revit__.ViewActivated -= self._view_activated_handler
                    self._view_activated_handler = None
            except:
                pass
            # Close the window
            window = self.Parent
            while window and not hasattr(window, 'Close'):
                window = window.Parent
            if window and hasattr(window, 'Close'):
                window.Close()
        except Exception as ex:
            print("Error closing window: {}".format(str(ex)))

def show_window():
    try:
        handler = ExternalEventHandler()
        external_event = ExternalEvent.Create(handler)

        view_model = MainViewModel(external_event, handler)
        handler.view_model = view_model

        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, "UI.xaml")
        window = IsolateWindow(xaml_path, view_model)
        window.Show()
    except Exception as ex:
        print("Error showing window: {}".format(str(ex)))
        log_error()


if __name__ == '__main__':
    show_window()