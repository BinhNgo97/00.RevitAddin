# -*- coding: utf-8 -*-

from pyrevit import revit, forms
from Autodesk.Revit.UI import ExternalEvent
from Autodesk.Revit.DB import FilteredElementCollector, Family
from Main_Handler import ExternalEventHandler
from Main_ViewModel import VoidCreationViewModel
import os
import sys
import traceback

doc = revit.doc
_modeless_window = None


class CreateVoidWindow(forms.WPFWindow):
    def __init__(self, xaml_file_path):
        forms.WPFWindow.__init__(self, xaml_file_path)
        
        # Initialize Handler first
        self._handler = ExternalEventHandler()
        self._ext = ExternalEvent.Create(self._handler)
        
        # Register completion callback
        self._handler.OnCompleted = self._on_handler_completed
        
        # Initialize ViewModel
        self.ViewModel = VoidCreationViewModel()
        
        # Set DataContext for binding
        self.DataContext = self.ViewModel
        
        # Wire button events
        self.btnSelectElements.Click += self._on_select_elements
        self.btnCreate.Click += self._on_create_void
        
        # Initialize: Generate unique default name
        self._generate_unique_name()
    
    def _generate_unique_name(self):
        """Generate a unique family name that doesn't exist in project"""
        default_name = "Void_From_SolidUnion"
        try:
            if doc is None:
                self.ViewModel.FamilyName = default_name
                return
                
            existing_families = [f.Name for f in FilteredElementCollector(doc).OfClass(Family)]
            fam_name = default_name
            counter = 1
            while fam_name in existing_families:
                fam_name = "{}_{}".format(default_name, counter)
                counter += 1
            self.ViewModel.FamilyName = fam_name
        except Exception as ex:
            # print("Generate unique name error: {}".format(str(ex)))
            self.ViewModel.FamilyName = default_name
    
    def _on_select_elements(self, sender, e):
        """Handle Select Elements button click"""
        try:
            if self._handler.is_busy:
                forms.alert("Please wait, operation in progress...")
                return
            
            self.ViewModel.StatusMessage = "Selecting elements..."
            self._handler.action = "select_elements"
            self._ext.Raise()
        except Exception as ex:
            print("Select elements error: {}".format(str(ex)))
            self.ViewModel.StatusMessage = "Selection error: {}".format(str(ex))
    
    def _on_create_void(self, sender, e):
        """Handle Create button click"""
        try:
            if self._handler.is_busy:
                forms.alert("Please wait, operation in progress...")
                return
            
            if not self._handler.SelectedElements or len(self._handler.SelectedElements) == 0:
                forms.alert("Please select elements first!")
                return
            
            if not self.ViewModel.FamilyName or not self.ViewModel.FamilyName.strip():
                forms.alert("Please enter a family name!")
                return
            
            fam_name = self.ViewModel.FamilyName.strip()
            
            # Don't check family existence here - let Handler do it in Revit API context
            # Just set the confirm flag to True by default
            self._handler.ConfirmOverwrite = True
            
            self.ViewModel.StatusMessage = "Creating void family..."
            self._handler.FamilyName = fam_name
            self._handler.action = "create_void"
            self._ext.Raise()
            
        except Exception as ex:
            print("Create void error: {}".format(str(ex)))
            import traceback as tb
            tb.print_exc()
            self.ViewModel.StatusMessage = "Create error: {}".format(str(ex))
    
    def _on_handler_completed(self, success, action_name):
        """Callback when external event completes"""
        try:
            if action_name == "select_elements":
                if success:
                    count = len(self._handler.SelectedElements)
                    self.ViewModel.SelectedCount = count
                    # Don't update UI during selection - only store count internally
                else:
                    self.ViewModel.SelectedCount = 0
                    
            elif action_name == "create_void":
                if success:
                    # Update result display
                    result_text = "✅ VOID FAMILY CREATED\n\n"
                    result_text += "Family Name: {}\n\n".format(self._handler.CreatedFamilyName)
                    result_text += "⚠️ IMPORTANT:\n"
                    result_text += "This is a FACE-BASED family.\n\n"
                    result_text += "To cut geometry:\n"
                    result_text += "1. Select the void instance\n"
                    result_text += "2. Place it on a wall/floor face\n"
                    result_text += "3. Use 'Cut Geometry' tool"
                    
                    self.ViewModel.StatusMessage = result_text
                    self.ViewModel.CreatedFamilyName = self._handler.CreatedFamilyName
                    
                    # Reset for next operation
                    self._handler.SelectedElements = []
                    self.ViewModel.SelectedCount = 0
                    self._generate_unique_name()
                else:
                    self.ViewModel.StatusMessage = "❌ FAILED\n\n{}".format(self._handler.message)
                    
        except Exception as ex:
            print("OnCompleted error: {}".format(str(ex)))
            import traceback as tb
            tb.print_exc()


def show_window():
    global _modeless_window
    
    # Close existing window if any
    if _modeless_window:
        try:
            _modeless_window.Close()
        except:
            pass
        _modeless_window = None
    
    # Get XAML path
    xaml_path = os.path.join(os.path.dirname(__file__), "UI.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("UI.xaml not found:\n{}".format(xaml_path))
        return
    
    try:
        # Create and show window
        win = CreateVoidWindow(xaml_path)
        
        def _on_closed(sender, e):
            global _modeless_window
            _modeless_window = None
        
        win.Closed += _on_closed
        win.Show()
        _modeless_window = win
        
    except Exception as ex:
        import traceback as tb
        tb.print_exc()
        forms.alert("Error showing Create Void window:\n{}".format(str(ex)))


def main():
    show_window()

if __name__ == "__main__":
    main()
