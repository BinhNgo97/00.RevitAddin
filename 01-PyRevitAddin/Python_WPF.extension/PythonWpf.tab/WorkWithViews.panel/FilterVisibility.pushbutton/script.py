# -*- coding: utf-8 -*-

import os
import clr
clr.AddReference("WindowsBase")
from pyrevit import forms
from Autodesk.Revit.UI import ExternalEvent
from System import Action
from System.Windows.Threading import Dispatcher, DispatcherPriority

from Main_EventHandler import ExternalEventHandler
from Main_ViewModel import MainViewModel

class FilterVisibilityWindow(forms.WPFWindow):
    def __init__(self, xaml_path, external_event, handler, view_model):
        forms.WPFWindow.__init__(self, xaml_path)
        self._ext = external_event
        self._handler = handler
        self._vm = view_model

        self.DataContext = self._vm

        # Load initial filters after window is shown
        # Use a small delay to ensure window is fully initialized
        def load_filters_delayed():
            try:
                self._vm.load_filters()
                # No need for sleep anymore, refresh handled in handler
            except Exception as ex:
                print("Error in load_filters_delayed: {}".format(str(ex)))
        
        # Schedule delayed load
        Dispatcher.CurrentDispatcher.BeginInvoke(
            Action(load_filters_delayed),
            DispatcherPriority.Loaded
        )
    
    def refresh_filters_ui(self):
        """Refresh the ListBox to show updated filters"""
        try:
            # Force UI refresh by re-binding
            if hasattr(self, 'lstFilters'):
                self.lstFilters.ItemsSource = None
                self.lstFilters.ItemsSource = self._vm.Filters
                print("UI refreshed: {} filters displayed".format(self.lstFilters.Items.Count))
        except Exception as ex:
            print("Error refreshing UI: {}".format(str(ex)))

    def CheckBox_Changed(self, sender, e):
        """Handler for CheckBox Checked/Unchecked events"""
        try:
            # Prevent rapid clicks from causing issues
            if self._handler.is_busy or self._handler._is_executing:
                return
            self._vm.apply_visibility()
        except:
            # Don't crash if handler fails
            pass


def show_window():
    handler = ExternalEventHandler()
    ext = ExternalEvent.Create(handler)

    vm = MainViewModel(ext, handler)
    handler.view_model = vm

    xaml_path = os.path.join(os.path.dirname(__file__), "UI.xaml")
    window = FilterVisibilityWindow(xaml_path, ext, handler, vm)
    handler.window = window  # Add this to reference window in handler
    window.Show()


if __name__ == "__main__":
    show_window()