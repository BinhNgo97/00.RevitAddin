# -*- coding: utf-8 -*-
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows.Media import SolidColorBrush, Color as MediaColor


class PendingChangeSet(object):
    """Buffer for visibility changes to commit intentionally."""
    def __init__(self):
        # store by integer id for stable hashing
        self._vis_changes = {}   # id(int) -> bool
        self._id_map = {}        # id(int) -> ElementId

    def set_visibility(self, filter_id, is_visible):
        key = getattr(filter_id, 'IntegerValue', None)
        key = key if key is not None else int(filter_id)
        self._vis_changes[key] = bool(is_visible)
        self._id_map[key] = filter_id

    def remove(self, filter_id):
        key = getattr(filter_id, 'IntegerValue', None)
        key = key if key is not None else int(filter_id)
        if key in self._vis_changes:
            del self._vis_changes[key]
        if key in self._id_map:
            del self._id_map[key]

    def clear(self):
        self._vis_changes.clear()
        self._id_map.clear()

    def has_changes(self):
        return len(self._vis_changes) > 0

    def items(self):
        # returns iterable of (ElementId, bool)
        return [(self._id_map[k], self._vis_changes[k]) for k in list(self._vis_changes.keys())]

class FilterItem(INotifyPropertyChanged):
    """ViewModel for a single filter item"""
    
    def __init__(self, filter_id, name, is_enabled):
        self._filter_id = filter_id
        self._name = name
        self._is_enabled = is_enabled
        self._color_brush = None
        self._property_changed_handlers = []
    
    # INotifyPropertyChanged implementation
    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    
    def OnPropertyChanged(self, prop_name):
        """Raise property changed event"""
        args = PropertyChangedEventArgs(prop_name)
        # Debug logging (optional - uncomment to debug)
        # print("Property changed: {} = {}".format(prop_name, getattr(self, prop_name)))
        for handler in self._property_changed_handlers:
            handler(self, args)
    
    # Properties
    @property
    def FilterId(self):
        return self._filter_id
    
    @property
    def Name(self):
        return self._name
    
    @property
    def IsEnabled(self):
        return self._is_enabled
    
    @IsEnabled.setter
    def IsEnabled(self, value):
        if self._is_enabled != value:
            self._is_enabled = value
            self.OnPropertyChanged("IsEnabled")

    # Color brush for UI binding (button background)
    @property
    def ColorBrush(self):
        return self._color_brush

    @ColorBrush.setter
    def ColorBrush(self, brush):
        if self._color_brush is not brush:
            self._color_brush = brush
            self.OnPropertyChanged("ColorBrush")

    def SetColorRGB(self, r, g, b):
        try:
            brush = SolidColorBrush(MediaColor.FromRgb(r, g, b))
            self.ColorBrush = brush
        except Exception:
            self.ColorBrush = None
