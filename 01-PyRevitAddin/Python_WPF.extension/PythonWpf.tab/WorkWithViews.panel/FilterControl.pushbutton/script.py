class dotCamera_t(object):
 # no doc
 DirectionVector=None
 FieldOfView=None
 Location=None
 UpVector=None
 ZoomFactor=None
import clr
clr.AddReference('System.Windows.Forms')
# Add parent directory to Python path
# sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('System.Core') 
# For ObservableCollection
from pyrevit import forms, revit
from Autodesk.Revit.UI import ExternalEvent
from Autodesk.Revit.DB import FilteredElementCollector, ParameterFilterElement
# import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
# import System
# from System.Windows.Forms import ColorDialog, DialogResult
# from System import Action
from Main_Handler import ExternalEventHandler
import Main_ViewModel as vm

# from System.Windows.Threading import DispatcherPriority


doc = revit.doc

_modeless_window = None


class FilterControlWindow(forms.WPFWindow):
    def __init__(self, xaml):
        forms.WPFWindow.__init__(self, xaml)

        self._handler = ExternalEventHandler()
        # Always initialize with the current active view
        try:
            self._handler.View = revit.doc.ActiveView
        except:
            self._handler.View = None
        self._ext = ExternalEvent.Create(self._handler)

        # Register completion callback (no polling)
        self._handler.OnCompleted = self._on_handler_completed
        # Bindable collections
        self.ProjectFilters = ObservableCollection[vm.FilterItem]()
        self.ViewFilters = ObservableCollection[vm.FilterItem]()
        self.Pending = vm.PendingChangeSet()
        self.DataContext = self

        # State
        self._clear_pending_on_success = False

        # Wire buttons
        self.btnRefresh.Click += self._on_refresh
        self.btnApplyVisibility.Click += self._on_apply_visibility
        self.btnSetColorRandom.Click += self._on_set_color_random
        self.btnAdd.Click += self._on_add_filters
        self.btnRemove.Click += self._on_remove_filters
        # Move up/down controls for View list
        self.btnMoveUp.Click += self._on_move_up
        self.btnMoveDown.Click += self._on_move_down

        # Load initial data
        # NOTE: Initial refresh when window opens - loads filters for current active view
        self._load_collections()
        
        # TODO: Add view change detection here if needed
        # Example implementation with polling timer:
        # from System.Windows.Threading import DispatcherTimer, DispatcherPriority
        # self._last_view_id = None
        # self._view_check_timer = DispatcherTimer()
        # self._view_check_timer.Interval = TimeSpan.FromSeconds(1)
        # self._view_check_timer.Tick += self._check_view_changed
        # self._view_check_timer.Start()

    def _load_collections(self):
        """Refresh data in both lstProjectFilters and lstViewFilters.
        
        NOTE: REFRESH MECHANISM
        =======================
        This method is called automatically in these scenarios:
        1. When window is first opened (in __init__)
        2. After add_filters operation completes (via _on_handler_completed callback)
        3. After remove_filters operation completes (via _on_handler_completed callback)
        4. After apply_visibility operation completes (via _on_handler_completed callback)
        5. After set_filter_colors operation completes (via _on_handler_completed callback)
        
        KNOWN LIMITATION - View Change Detection:
        ==========================================
        Currently does NOT automatically refresh when user switches to a different view in Revit.
        The active view is read fresh each time this method runs, but the method is not
        triggered by view changes.
        
        SOLUTION OPTIONS for View Change Detection:
        ============================================
        Option 1: Add manual Refresh button in UI
          - Simple, reliable
          - Requires user action
          
        Option 2: Polling timer
          - Add DispatcherTimer to check if active view changed every 1-2 seconds
          - Store self._last_view_id and compare on timer tick
          - Auto-refresh when different
          
        Option 3: Revit ViewActivated event (if available in pyRevit)
          - Most elegant but may not be available in modeless context
        """
        try:
            # Resolve a fresh active view every reload to avoid stale references
            # Use global doc variable (defined at module level)
            from pyrevit import revit as rvt
            from Autodesk.Revit.DB import FilteredElementCollector, ParameterFilterElement
            import Main_ViewModel as vm
            current_view = None
            try:
                current_view = rvt.doc.ActiveView
            except Exception as ex:
                print("_load_collections: Error getting ActiveView - {}".format(str(ex)))
                current_view = None

            if current_view is None:
                print("_load_collections: No active view - skipping refresh")
                return
                
            if getattr(current_view, 'IsTemplate', False):
                # Do not modify lists if no valid view
                print("_load_collections: Active view is a template - skipping refresh")
                return
            
            print("_load_collections: Processing view '{}'".format(getattr(current_view, 'Name', 'Unknown')))

            # Build new lists first; only swap into UI on success
            new_view_items = []
            new_project_items = []

            try:
                view_filter_ids = set([fid for fid in current_view.GetFilters()])
                print("_load_collections: Found {} filters in view".format(len(view_filter_ids)))
            except Exception as ex:
                # If API returns nothing, keep current UI state
                print("_load_collections: Error getting view filters - {}".format(str(ex)))
                return

            all_filters = FilteredElementCollector(rvt.doc).OfClass(ParameterFilterElement).ToElements()
            print("_load_collections: Found {} total filters in project".format(len(list(all_filters))))
            for f in all_filters:
                fid = f.Id
                name = f.Name
                if fid in view_filter_ids:
                    try:
                        enabled = current_view.GetFilterVisibility(fid)
                    except:
                        enabled = True
                    item = vm.FilterItem(fid, name, enabled)
                    item.add_PropertyChanged(self._on_view_item_changed)
                    # Try to read existing override color if API supports getters
                    try:
                        ogs = current_view.GetFilterOverrides(fid)
                        if ogs is not None:
                            color_val = None
                            try:
                                color_val = ogs.GetSurfaceForegroundPatternColor()
                            except Exception:
                                try:
                                    color_val = ogs.GetProjectionFillPatternColor()
                                except Exception:
                                    try:
                                        color_val = ogs.GetCutForegroundPatternColor()
                                    except Exception:
                                        try:
                                            color_val = ogs.GetCutFillPatternColor()
                                        except Exception:
                                            color_val = None
                            if color_val is not None:
                                try:
                                    item.SetColorRGB(color_val.Red, color_val.Green, color_val.Blue)
                                except Exception:
                                    pass
                    except Exception:
                        pass
                    new_view_items.append(item)
                else:
                    new_project_items.append(vm.FilterItem(fid, name, False))

            # If we reached here, we have consistent snapshots; now update UI collections
            # NOTE: Clear and rebuild both collections - this updates lstViewFilters
            print("_load_collections: Updating UI - {} view items, {} project items".format(
                len(new_view_items), len(new_project_items)))
            
            self.ViewFilters.Clear()
            for it in new_view_items:
                self.ViewFilters.Add(it)

            # NOTE: Clear and rebuild project filters - this updates lstProjectFilters  
            self.ProjectFilters.Clear()
            for it in new_project_items:
                self.ProjectFilters.Add(it)
                
            print("_load_collections: UI updated successfully")
        except Exception as ex:
            import traceback
            print("Error loading filters: {}".format(str(ex)))
            traceback.print_exc()

    # Selection capture and batch apply at UI layer (Level 3)
    def OnViewCheckBoxChecked(self, sender, e):
        try:
            selected = self._get_selected_from("lstViewFilters")
            # If multiple selected, UI drives batch by setting properties directly
            if selected and len(selected) > 1:
                for it in selected:
                    it.IsEnabled = True
        except Exception as ex:
            print("OnViewCheckBoxChecked error: {}".format(str(ex)))

    def OnViewCheckBoxUnchecked(self, sender, e):
        try:
            selected = self._get_selected_from("lstViewFilters")
            if selected and len(selected) > 1:
                for it in selected:
                    it.IsEnabled = False
        except Exception as ex:
            print("OnViewCheckBoxUnchecked error: {}".format(str(ex)))

    def _get_selected_from(self, list_name):
        selected = []
        try:
            lb = self.FindName(list_name)
            if lb:
                for item in lb.SelectedItems:
                    selected.append(item)
        except Exception as ex:
            print("Get selected items error ({}): {}".format(list_name, str(ex)))
        return selected

    def _on_view_item_changed(self, sender, e):
        # Pure VM: only record pending for the changed item
        if not sender or e.PropertyName != "IsEnabled":
            return
        self.Pending.set_visibility(sender.FilterId, sender.IsEnabled)

    # Buttons
    def _on_refresh(self, sender, e):
        """Manual refresh - reload filters from current active view.
        
        NOTE: Use this when:
        - Switching to a different view in Revit
        - Creating new filters in Revit
        - Any time lists appear out of sync
        """
        try:
            # Update handler's view reference to current active view
            from pyrevit import revit as rvt
            try:
                current_view = rvt.doc.ActiveView
                self._handler.View = current_view
                print("Active view: {} (IsTemplate: {})".format(
                    getattr(current_view, 'Name', 'Unknown'),
                    getattr(current_view, 'IsTemplate', 'Unknown')
                ))
            except Exception as view_ex:
                self._handler.View = None
                print("Error getting active view: {}".format(str(view_ex)))
                return
            
            # Reload both lists from Revit
            print("Loading collections...")
            self._load_collections()
            print("Filters refreshed - ViewFilters: {}, ProjectFilters: {}".format(
                self.ViewFilters.Count if hasattr(self.ViewFilters, 'Count') else len(self.ViewFilters),
                self.ProjectFilters.Count if hasattr(self.ProjectFilters, 'Count') else len(self.ProjectFilters)
            ))
        except Exception as ex:
            import traceback
            print("Refresh error: {}".format(str(ex)))
            traceback.print_exc()
    
    def _on_apply_visibility(self, sender, e):
        try:
            if not self.Pending.has_changes():
                return
            self._handler.VisibilityChanges = self.Pending.items()
            self._handler.action = "apply_visibility"
            self._ext.Raise()
            # Defer clearing pending until success; OnCompleted will handle refresh
            self._clear_pending_on_success = True
        except Exception as ex:
            print("Apply visibility error: {}".format(str(ex)))

    def OnPickColorClick(self, sender, e):
        """Open a color picker and set override for the clicked filter."""
        try:
            # sender is the Button inside the row; its DataContext is FilterItem
            item = getattr(sender, 'DataContext', None)
            if item is None:
                return
            # Use Windows Forms ColorDialog for simplicity
            from System.Windows.Forms import ColorDialog, DialogResult
            from Autodesk.Revit.DB import Color
            cd = ColorDialog()
            if cd.ShowDialog() == DialogResult.OK:
                c = cd.Color
                dbcolor = Color(c.R, c.G, c.B)
                self._handler.ColorOverrides = [(item.FilterId, dbcolor)]
                self._handler.action = "set_filter_colors"
                # Update UI brush immediately
                try:
                    item.SetColorRGB(c.R, c.G, c.B)
                except Exception as uiex:
                    print("Update item brush error: {}".format(uiex))
                self._ext.Raise()
        except Exception as ex:
            print("Pick color error: {}".format(ex))

    def _on_add_filters(self, sender, e):
        """Add selected filters from project list to current view.
        
        NOTE: REFRESH BEHAVIOR
        =======================
        1. Optimistic UI update: Immediately moves items between lists for responsiveness
        2. ExternalEvent raised: Actual Revit transaction happens asynchronously
        3. _on_handler_completed: Called after transaction, triggers _load_collections()
        4. Authoritative refresh: Both lists rebuilt from Revit data, ensuring accuracy
        """
        try:
            selected = self._get_selected_from("lstProjectFilters")
            if not selected:
                return
            # Optimistically move items to the view list for immediate UI feedback
            for it in list(selected):
                try:
                    self.ProjectFilters.Remove(it)
                except Exception as rem_ex:
                    print("Remove from ProjectFilters error: {}".format(rem_ex))
                try:
                    it.add_PropertyChanged(self._on_view_item_changed)
                except Exception as hook_ex:
                    print("Add PropertyChanged hook error: {}".format(hook_ex))
                try:
                    self.ViewFilters.Add(it)
                except Exception as add_ex:
                    print("Add to ViewFilters error: {}".format(add_ex))

            # Raise ExternalEvent to add filters in Revit
            self._handler.FilterIds = [it.FilterId for it in selected]
            self._handler.action = "add_filters"
            self._ext.Raise()
            # OnCompleted will handle authoritative refresh
        except Exception as ex:
            print("Add filters error: {}".format(str(ex)))

    def _on_remove_filters(self, sender, e):
        """Remove selected filters from current view.
        
        NOTE: REFRESH BEHAVIOR  
        =======================
        Same refresh pattern as add_filters:
        1. Optimistic UI update
        2. ExternalEvent for Revit transaction
        3. Authoritative refresh via _load_collections() after completion
        """
        try:
            selected = self._get_selected_from("lstViewFilters")
            if not selected:
                return
            # Optimistically move items back to project list for immediate UI feedback
            for it in list(selected):
                try:
                    self.ViewFilters.Remove(it)
                except Exception as rem_ex:
                    print("Remove from ViewFilters error: {}".format(rem_ex))
                try:
                    # Stop recording visibility on non-view items
                    it.remove_PropertyChanged(self._on_view_item_changed)
                except Exception as unhook_ex:
                    print("Remove PropertyChanged hook error: {}".format(unhook_ex))
                try:
                    self.Pending.remove(it.FilterId)
                except Exception as pend_ex:
                    print("Pending remove error: {}".format(pend_ex))
                try:
                    # Reuse the existing item, mark not enabled for project list
                    try:
                        it.IsEnabled = False
                    except Exception:
                        pass
                    self.ProjectFilters.Add(it)
                except Exception as add_ex:
                    print("Add to ProjectFilters error: {}".format(add_ex))

            # Raise ExternalEvent to remove filters in Revit
            self._handler.FilterIds = [it.FilterId for it in selected]
            self._handler.action = "remove_filters"
            self._ext.Raise()
            # OnCompleted will handle authoritative refresh
        except Exception as ex:
            print("Remove filters error: {}".format(str(ex)))

    def _on_set_color_random(self, sender, e):
        """Assign random colors to selected view filters only."""
        try:
            selected = self._get_selected_from("lstViewFilters")
            if not selected:
                # nothing selected; do nothing per requirement 2.1
                return
            overrides = []
            for it in selected:
                import random
                from Autodesk.Revit.DB import Color
                r = random.randint(0, 255)
                g = random.randint(0, 255)
                b = random.randint(0, 255)
                overrides.append((it.FilterId, Color(r, g, b)))
                try:
                    it.SetColorRGB(r, g, b)
                except Exception:
                    pass
            self._handler.ColorOverrides = overrides
            self._handler.action = "set_filter_colors"
            self._ext.Raise()
        except Exception as ex:
            print("Set color random error: {}".format(ex))

    def _on_move_up(self, sender, e):
        try:
            lv = self.FindName("lstViewFilters")
            if lv is None:
                return
            items = self.ViewFilters
            count = items.Count if hasattr(items, 'Count') else len(items)
            if count < 2:
                return
            selected_items = [it for it in lv.SelectedItems]
            if not selected_items:
                return
            # Indices of selected items
            try:
                indices = sorted([items.IndexOf(it) for it in selected_items])
            except Exception:
                raw = []
                for it in selected_items:
                    for idx in range(count):
                        if items[idx] is it:
                            raw.append(idx)
                            break
                indices = sorted(raw)

            moved = False
            for idx in indices:
                if idx > 0 and (idx - 1) not in indices:
                    it = items[idx]
                    items.RemoveAt(idx)
                    items.Insert(idx - 1, it)
                    moved = True
            if moved:
                try:
                    lv.SelectedItems.Clear()
                except Exception:
                    pass
                for it in selected_items:
                    lv.SelectedItems.Add(it)
                try:
                    lv.ScrollIntoView(selected_items[0])
                except Exception:
                    pass
        except Exception as ex:
            print("Move up error: {}".format(ex))

    def _on_move_down(self, sender, e):
        try:
            lv = self.FindName("lstViewFilters")
            if lv is None:
                return
            items = self.ViewFilters
            count = items.Count if hasattr(items, 'Count') else len(items)
            if count < 2:
                return
            selected_items = [it for it in lv.SelectedItems]
            if not selected_items:
                return
            # Indices of selected items (process from bottom)
            try:
                indices = sorted([items.IndexOf(it) for it in selected_items], reverse=True)
            except Exception:
                raw = []
                for it in selected_items:
                    for idx in range(count):
                        if items[idx] is it:
                            raw.append(idx)
                            break
                indices = sorted(raw, reverse=True)

            moved = False
            for idx in indices:
                if idx < count - 1 and (idx + 1) not in indices:
                    it = items[idx]
                    items.RemoveAt(idx)
                    items.Insert(idx + 1, it)
                    moved = True
            if moved:
                try:
                    lv.SelectedItems.Clear()
                except Exception:
                    pass
                for it in selected_items:
                    lv.SelectedItems.Add(it)
                try:
                    lv.ScrollIntoView(selected_items[-1])
                except Exception:
                    pass
        except Exception as ex:
            print("Move down error: {}".format(ex))

    # def _refresh_after_event(self):
    def _on_handler_completed(self, success, action_name):
        """Callback triggered after ExternalEvent operations complete.
        
        NOTE: AUTO-REFRESH AFTER OPERATIONS
        ====================================
        This is the main refresh trigger that handles:
        - After add_filters: Both lists refresh to show filter moved from Project to View
        - After remove_filters: Both lists refresh to show filter moved from View to Project  
        - After apply_visibility: View list refreshes to show updated visibility states
        - After set_filter_colors: View list refreshes to show updated colors
        - After any new filter created in Revit: If user re-opens window, new filters appear
        
        The _load_collections() call below re-reads ALL filters from Revit and rebuilds
        both lstProjectFilters and lstViewFilters from scratch.
        """
        try:
            # Refresh immediately after transaction commit
            # NOTE: This ensures UI is synchronized with Revit after any filter operation
            self._load_collections()
            if self._clear_pending_on_success and success:
                self.Pending.clear()
        except Exception as ex:
            print("OnCompleted refresh error: {}".format(str(ex)))
        finally:
            self._clear_pending_on_success = False

def show_window():
    global _modeless_window
    # Resolve fresh active view at show time
    try:
        current_view = revit.doc.ActiveView
    except:
        current_view = None
    if not current_view or current_view.IsTemplate:
        forms.alert("Please open a valid view (not a template) before running this command.")
        return
    if _modeless_window:
        try:
            _modeless_window.Close()
        except:
            pass
        _modeless_window = None
    xaml_path = os.path.join(os.path.dirname(__file__), "FilterWindow.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("FilterWindow.xaml not found:\n{}".format(xaml_path))
        return
    try:
        win = FilterControlWindow(xaml_path)
        def _on_closed(sender, e):
            global _modeless_window
            _modeless_window = None
        win.Closed += _on_closed
        win.Show()
        _modeless_window = win
    except Exception as ex:
        import traceback
        traceback.print_exc()
        forms.alert("Error showing Filter Control window:\n{}".format(str(ex)))


def main():
    show_window()
    
if __name__ == "__main__":
    main()
