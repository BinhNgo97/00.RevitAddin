# -*- coding: utf-8 -*-
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs


class VoidCreationViewModel(INotifyPropertyChanged):
    """ViewModel for Create Void from Link window"""
    
    def __init__(self):
        self._family_name = "Void_From_SolidUnion"
        self._selected_count = 0
        self._created_family_name = ""
        self._status_message = "Ready"
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
        for handler in self._property_changed_handlers:
            handler(self, args)
    
    # Properties
    @property
    def FamilyName(self):
        return self._family_name
    
    @FamilyName.setter
    def FamilyName(self, value):
        if self._family_name != value:
            self._family_name = value
            self.OnPropertyChanged("FamilyName")
    
    @property
    def SelectedCount(self):
        return self._selected_count
    
    @SelectedCount.setter
    def SelectedCount(self, value):
        if self._selected_count != value:
            self._selected_count = value
            self.OnPropertyChanged("SelectedCount")
            self.OnPropertyChanged("SelectionStatus")
    
    @property
    def SelectionStatus(self):
        if self._selected_count > 0:
            return "{} element(s) selected".format(self._selected_count)
        return "No elements selected"
    
    @property
    def CreatedFamilyName(self):
        return self._created_family_name
    
    @CreatedFamilyName.setter
    def CreatedFamilyName(self, value):
        if self._created_family_name != value:
            self._created_family_name = value
            self.OnPropertyChanged("CreatedFamilyName")
    
    @property
    def StatusMessage(self):
        return self._status_message
    
    @StatusMessage.setter
    def StatusMessage(self, value):
        if self._status_message != value:
            self._status_message = value
            self.OnPropertyChanged("StatusMessage")
