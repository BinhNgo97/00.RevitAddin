# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import Transaction, FilteredElementCollector, ParameterFilterElement
from Autodesk.Revit.UI import IExternalEventHandler, TaskDialog

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
clr.AddReference("System.ComponentModel")

class FilterRow(INotifyPropertyChanged):
    __notify__ = None  # PropertyChanged event

    def __init__(self, name, fid, visible):
        self._name = name
        self._id = fid
        self._is_visible = visible

    @property
    def Name(self):
        return self._name

    @property
    def Id(self):
        return self._id

    @property
    def IsVisible(self):
        return self._is_visible

    @IsVisible.setter
    def IsVisible(self, value):
        if value != self._is_visible:
            self._is_visible = value
            self.OnPropertyChanged("IsVisible")

    def OnPropertyChanged(self, prop_name):
        if self.__notify__ is not None:
            self.__notify__(self, PropertyChangedEventArgs(prop_name))

    @property
    def PropertyChanged(self):
        return self.__notify__

    @PropertyChanged.setter
    def PropertyChanged(self, value):
        self.__notify__ = value


class ExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        self.action = None
        self.selected_filter_ids = []
        self.filters = []
        self.message = ""
        self.is_busy = False
        self._is_executing = False  # Guard to prevent re-entrancy
        self.view_model = None     # injected later
        self.window = None         # injected later

    def GetName(self):
        return "Filter Visibility Handler"

    def Execute(self, uiapp):
        # Prevent re-entrancy - critical for stability
        if self._is_executing:
            return
        
        self._is_executing = True
        self.is_busy = True
        
        try:
            # Validate uiapp
            if not uiapp:
                return
                
            # Validate active document using pyRevit global for safety
            try:
                uidoc = __revit__.ActiveUIDocument  # Use __revit__ for reliable access
                doc = uidoc.Document if uidoc else None
                if not doc or doc.IsClosed:
                    return
                print("DEBUG: Using doc Title: {}".format(doc.Title))  # Debug to confirm doc
            except:
                return

            if self.action == "load_filters":
                self.filters = self._load_filters(uiapp)
                # Debug: print filter count
                filter_count = len(self.filters) if self.filters else 0
                print("Loaded {} filters from view".format(filter_count))
                
                # Update ViewModel safely - this triggers UI update
                if self.view_model and self.filters is not None:
                    try:
                        # Update the ObservableCollection
                        self.view_model.Filters = self.filters
                        print("ViewModel.Filters updated: {} items in collection".format(self.view_model.Filters.Count))
                        
                        # Force refresh UI from main thread
                        from System.Windows.Threading import Dispatcher
                        from System import Action
                        
                        def refresh_ui():
                            try:
                                if self.window:
                                    self.window.refresh_filters_ui()
                            except:
                                pass
                        
                        Dispatcher.CurrentDispatcher.Invoke(
                            Action(refresh_ui),
                            DispatcherPriority.Background
                        )
                    except Exception as ex:
                        # Print error for debugging
                        print("Error setting ViewModel.Filters: {}".format(str(ex)))
                        import traceback
                        traceback.print_exc()
                        pass
                else:
                    print("Warning: view_model is None or filters is None")

            elif self.action == "set_visibility":
                self.message = self._set_visibility(uiapp)
                # Always show message
                try:
                    TaskDialog.Show("Filter Visibility", self.message)
                except:
                    pass
                        
        except Exception as ex:
            # Catch all exceptions to prevent crash
            try:
                self.message = "Error: {}".format(str(ex))
                TaskDialog.Show("Filter Visibility Error", self.message)
            except:
                # Even TaskDialog can fail, so catch that too
                pass
        finally:
            self._is_executing = False
            self.is_busy = False
            self.action = None

    # -----------------------
    # Load filters
    # -----------------------
    def _load_filters(self, uiapp):
        """Load all filters in the document, check if applied to the active view."""
        try:
            uidoc = __revit__.ActiveUIDocument  # Use __revit__ for safety
            doc = uidoc.Document if uidoc else None
            view = uidoc.ActiveView if uidoc else None
            if not doc or doc.IsClosed or not view:
                return []

            filters = []
            # Load all ParameterFilterElement in document
            raw_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
            print("DEBUG: Raw filters count: {}".format(len(raw_filters)))  # Debug to confirm collector

            for fe in raw_filters:
                try:
                    fid = fe.Id
                    if view.IsFilterApplied(fid):
                        visible = bool(view.GetFilterVisibility(fid))
                    else:
                        visible = False
                    print("CHECK: Adding filter - Name: {}, Id: {}, Visible: {}".format(fe.Name, fid.IntegerValue, visible))
                    filters.append(FilterRow(fe.Name, fid.IntegerValue, visible))
                except:
                    continue

            return sorted(filters, key=lambda x: x.Name)
        except:
            return []


    # -----------------------
    # Set visibility
    # -----------------------
    def _set_visibility(self, uiapp):
        try:
            uidoc = __revit__.ActiveUIDocument  # Use __revit__ for safety
            if not uidoc:
                return "No active document."

            doc = uidoc.Document
            if not doc or doc.IsClosed:
                return "Document is closed or invalid."

            view = uidoc.ActiveView
            if not view:
                return "No active view."

            selected_ids = set(self.selected_filter_ids) if self.selected_filter_ids else set()

            # Load all ParameterFilterElement again for setting
            raw_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
            if not raw_filters:
                return "No filters in the document."

            t = Transaction(doc, "Apply Filter Visibility")
            updated_count = 0

            try:
                t.Start()
                
                for fe in raw_filters:
                    # Validate ElementId
                    fid = fe.Id
                    if not fid or not hasattr(fid, 'IntegerValue'):
                        continue
                        
                    try:
                        fid_int = int(fid.IntegerValue)
                        should_be_visible = fid_int in selected_ids
                        
                        if should_be_visible:
                            if not view.IsFilterApplied(fid):
                                view.AddFilter(fid)
                            view.SetFilterVisibility(fid, True)
                        else:
                            if view.IsFilterApplied(fid):
                                view.SetFilterVisibility(fid, False)
                            # Không remove, chỉ set invisible để giữ apply nếu cần
                        
                        updated_count += 1
                    except:
                        # Continue with other filters if one fails
                        continue
                        
                t.Commit()
                return "Updated {} filter(s) visibility.".format(updated_count)
            except Exception as ex:
                if t.HasStarted(): 
                    try:
                        t.RollBack()
                    except:
                        pass
                return "Failed: {}".format(str(ex))
        except Exception as ex:
            return "Error: {}".format(str(ex))