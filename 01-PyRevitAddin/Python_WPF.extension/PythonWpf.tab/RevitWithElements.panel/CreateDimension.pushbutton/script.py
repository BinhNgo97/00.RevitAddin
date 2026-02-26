# -*- coding: utf-8 -*-
import os, clr
clr.AddReference('PresentationFramework')
clr.AddReference('RevitAPIUI')

from pyrevit import forms
from Autodesk.Revit.UI import ExternalEvent

from Main_ViewModel import DimensionViewModel
from Main_EventHandler import MainEventHandler


class DimensionWindow(forms.WPFWindow):
    def __init__(self):
        xaml = os.path.join(os.path.dirname(__file__), "Dimension_UI.xaml")
        forms.WPFWindow.__init__(self, xaml)

        # ------------------------
        # ViewModel
        # ------------------------
        self.vm = DimensionViewModel()
        self.DataContext = self.vm

        # ------------------------
        # External Event
        # ------------------------
        self.handler = MainEventHandler(self, self.vm)
        self.ext_event = ExternalEvent.Create(self.handler)
        self.handler.OnCompleted = self._on_completed

    # -------- UI EVENTS --------
    def OnDimGridLevel(self, sender, args):
        self._run("GRID_LEVEL", "Selecting grids / levels...")

    def OnDimElementRefPlan(self, sender, args):
        self._run("ELEMENT_REFPLAN", "Selecting elements...")

    # -------- INTERNAL --------
    def _run(self, action, status):
        if self.handler.is_busy:
            return
        self.vm.StatusText = status
        self.handler.action = action
        self.ext_event.Raise()

    def _on_completed(self, success, action):
        self.vm.StatusText = self.handler.message or (
            "Done." if success else "Failed."
        )


def main():
    DimensionWindow().Show()


if __name__ == "__main__":
    main()
