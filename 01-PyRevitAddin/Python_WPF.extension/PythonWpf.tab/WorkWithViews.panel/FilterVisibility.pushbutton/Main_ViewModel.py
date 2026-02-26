# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
from System.Collections.ObjectModel import ObservableCollection


class MainViewModel(object):
    def __init__(self, external_event, handler):
        self._external_event = external_event
        self._handler = handler
        # Use ObservableCollection - it automatically notifies UI when items change
        # No need for PropertyChangedEventArgs!
        self._filters = ObservableCollection[object]()

    # -------------------------
    # Filters for binding - Using ObservableCollection
    # -------------------------
    @property
    def Filters(self):
        return self._filters

    @Filters.setter
    def Filters(self, value):
        """
        Update filters by clearing and adding new items.
        ObservableCollection will automatically notify UI of changes.
        """
        try:
            print("DEBUG: Filters setter called with value type: {}, count: {}".format(
                type(value), len(value) if value else 0))
            self._filters.Clear()
            print("DEBUG: Cleared ObservableCollection, current count: {}".format(self._filters.Count))
            if value:
                print("DEBUG: Adding {} items to ObservableCollection".format(len(value)))
                for idx, item in enumerate(value):
                    print("DEBUG: Adding item {}: {}".format(idx, item))
                    self._filters.Add(item)
                # Debug: print count to verify
                print("DEBUG: Filters updated: {} items in ObservableCollection".format(self._filters.Count))
            else:
                print("DEBUG: value is None or empty")
        except Exception as ex:
            # Print error for debugging
            print("Error updating Filters: {}".format(str(ex)))
            import traceback
            traceback.print_exc()
            pass

    # -------------------------
    # Commands to handler
    # -------------------------
    def load_filters(self):
        # Prevent multiple rapid calls
        if self._handler.is_busy or self._handler._is_executing:
            return
            
        try:
            self._handler.action = "load_filters"
            self._handler.is_busy = True
            self._external_event.Raise()
        except:
            # Don't crash if event raise fails
            pass

    def apply_visibility(self):
        # Prevent multiple rapid calls
        if self._handler.is_busy or self._handler._is_executing:
            return
            
        try:
            selected_ids = [f.Id for f in self._filters if f.IsVisible]

            self._handler.selected_filter_ids = selected_ids
            self._handler.action = "set_visibility"
            self._handler.is_busy = True
            self._external_event.Raise()
        except:
            # Don't crash if event raise fails
            pass