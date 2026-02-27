# -*- coding: utf-8 -*-
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs


class DimensionViewModel(INotifyPropertyChanged):
    def __init__(self):
        self._handlers = []
        self._status = "Ready."

    # -------- INotifyPropertyChanged --------
    def add_PropertyChanged(self, handler):
        if handler not in self._handlers:
            self._handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._handlers:
            self._handlers.remove(handler)

    def _notify(self, prop):
        args = PropertyChangedEventArgs(prop)
        for h in list(self._handlers):
            h(self, args)

    # -------- Bindable Properties --------
    @property
    def StatusText(self):
        return self._status

    @StatusText.setter
    def StatusText(self, value):
        self._status = value or ""
        self._notify("StatusText")
