# -*- coding: utf-8 -*-
# ===========================
# File: Main_ViewModel.py - Isolate Element By Parameter
# ===========================

import os, sys, clr
# Add parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

# Add Revit API references
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from aGeneral.ViewModel_Base import ViewModel_BaseEventHandler
from aGeneral.Command_Base import DelegateCommand
from System.Collections.ObjectModel import ObservableCollection
from Autodesk.Revit.DB import (
    FilteredElementCollector,
)
from System import String
from System.Windows.Forms import ColorDialog, DialogResult
from System.Windows.Media import Color

# pyRevit to get current doc/view safely from VM
from pyrevit import revit

class ParameterNameItem(ViewModel_BaseEventHandler):
    def __init__(self, name):
        ViewModel_BaseEventHandler.__init__(self)
        self._name = name
        self._is_selected = False

    @property
    def Name(self):
        return self._name

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected == bool(value):
            return
        self._is_selected = bool(value)
        self.OnPropertyChanged('IsSelected')


class ParameterRow(ViewModel_BaseEventHandler):
    def __init__(self, parameter_name, values):
        ViewModel_BaseEventHandler.__init__(self)
        self._parameter_name = parameter_name
        self._values = ObservableCollection[String]()
        for v in values:
            self._values.Add(String(v))
        self._selected_value = values[0] if len(values) > 0 else ""
        self._row_color = Color.FromRgb(200, 200, 200)
        # caches
        self._cached_ids = []           # ids from Get‑E
        self._last_override_ids = []    # ids last colored by this row
        self._has_user_color = False

    @property
    def ParameterName(self):
        return self._parameter_name

    @property
    def Values(self):
        return self._values

    @property
    def SelectedValue(self):
        return self._selected_value

    @SelectedValue.setter
    def SelectedValue(self, value):
        if self._selected_value == value:
            return
        self._selected_value = value
        # when value changes, clear cached ids (force re-match)
        self._cached_ids = []
        self.OnPropertyChanged('SelectedValue')
        self.OnPropertyChanged('CachedCount')

    @property
    def RowColor(self):
        return self._row_color

    def SetRowColor(self, clr):
        if clr is None:
            return
        if (self._row_color.R == clr.R and self._row_color.G == clr.G and self._row_color.B == clr.B):
            # still mark as user-set color
            self._has_user_color = True
            return
        self._row_color = clr
        self._has_user_color = True
        self.OnPropertyChanged('RowColor')

    # NEW: replace Values in-place and keep selection when possible
    def ReplaceValues(self, new_values, keep_selection=True):
        try:
            old_sel = self._selected_value
            self._values.Clear()
            for v in (new_values or []):
                self._values.Add(String(v))
            if keep_selection and old_sel in new_values:
                self._selected_value = old_sel
            else:
                self._selected_value = new_values[0] if new_values else ""
            # clear cache when options are rebuilt
            self._cached_ids = []
            self.OnPropertyChanged('Values')
            self.OnPropertyChanged('SelectedValue')
            self.OnPropertyChanged('CachedCount')
        except:
            pass

    # cache from Get‑E
    def SetCachedIds(self, ids_iterable):
        ints = []
        for i in (ids_iterable or []):
            try:
                ints.append(i.IntegerValue)
            except:
                try:
                    ints.append(int(i))
                except:
                    pass
        self._cached_ids = ints
        self.OnPropertyChanged('CachedCount')

    @property
    def CachedIds(self):
        return list(self._cached_ids)

    @property
    def CachedCount(self):
        try:
            return len(self._cached_ids)
        except:
            return 0

    # track last overridden ids so we can clear them when value changes
    def SetLastOverrideIds(self, ids_iterable):
        ints = []
        for i in (ids_iterable or []):
            try:
                ints.append(i.IntegerValue)
            except:
                try:
                    ints.append(int(i))
                except:
                    pass
        self._last_override_ids = ints

    @property
    def LastOverrideIds(self):
        return list(getattr(self, '_last_override_ids', []))

    @property
    def HasUserColor(self):
        return bool(self._has_user_color)


class MainViewModel(object):
    def __init__(self, external_event=None, handler=None):
        self.external_event = external_event
        self.handler = handler

    # simple proxy to raise select action (not used by simplified UI but kept for compatibility)
    def select_cad(self):
        if not self.handler or not self.external_event:
            return
        self.handler.action = "select_cad"
        self.external_event.Raise()
