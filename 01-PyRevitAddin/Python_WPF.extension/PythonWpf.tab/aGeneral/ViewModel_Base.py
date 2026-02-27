# -*- coding: utf-8 -*-
# base.py

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System import EventHandler

class ViewModel_BaseINotify(INotifyPropertyChanged):
    def __init__(self):
        # khởi tạo danh sách để lưu các handler của sự kiện PropertyChanged
        self._propertyChangedHandlers = []
        
    def add_PropertyChanged(self, handler):
        # thêm handler vào danh sách
        self._propertyChangedHandlers.append(handler)

    def remove_PropertyChanged(self, handler):
        # xoá handler khỏi danh sách nếu nó tồn tại
        if handler in self._propertyChangedHandlers:
            self._propertyChangedHandlers.remove(handler)

    def raise_property_changed(self, property_name):
        #gọi tất cả các handler trong danh sách với PropertyChangedEventArgs
        for handler in self._propertyChangedHandlers:
            handler(self, PropertyChangedEventArgs(property_name))


class ViewModel_BaseEventHandler(INotifyPropertyChanged):
    def __init__(self):
        self._propertyChangedHandlers = []
    
    def add_PropertyChanged(self, handler):
        self._propertyChangedHandlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        if handler in self._propertyChangedHandlers:
            self._propertyChangedHandlers.remove(handler)
    
    def OnPropertyChanged(self, propertyName):
        for handler in self._propertyChangedHandlers:
            handler(self, PropertyChangedEventArgs(propertyName))

