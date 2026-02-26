# -*- coding: utf-8 -*-
# File: Main_ViewModel.py - Duplicate Views

import sys
import os
import clr

# Add parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))


from aGeneral.ViewModel_Base import ViewModel_BaseEventHandler
from aGeneral.Command_Base import DelegateCommand
from System.Collections.ObjectModel import ObservableCollection
from System import String


class TemplateItem(ViewModel_BaseEventHandler):
    def __init__(self, element_id, name):
        ViewModel_BaseEventHandler.__init__(self)
        self._id = element_id
        self._name = name or ""

    @property
    def Id(self):
        return self._id

    @property
    def Name(self):
        return self._name


class ViewItem(ViewModel_BaseEventHandler):
    def __init__(self, element_id, name, template_name):
        ViewModel_BaseEventHandler.__init__(self)
        self._id = element_id
        self._name = name or ""
        self._template = template_name or "none"
        # per-row templates and selection
        self._templates = ObservableCollection[TemplateItem]()
        self._selected_template = None
        self._on_template_selected = None  # set by VM

    @property
    def ElementId(self):
        return self._id

    @property
    def ViewName(self):
        return self._name

    @ViewName.setter
    def ViewName(self, v):
        if self._name != v:
            self._name = v or ""
            self.OnPropertyChanged("ViewName")

    @property
    def ViewTemplate(self):
        return self._template

    @ViewTemplate.setter
    def ViewTemplate(self, v):
        if self._template != v:
            self._template = v or "none"
            self.OnPropertyChanged("ViewTemplate")

    @property
    def Templates(self):
        return self._templates

    @property
    def SelectedTemplate(self):
        return self._selected_template

    @SelectedTemplate.setter
    def SelectedTemplate(self, value):
        if self._selected_template is value:
            return
        self._selected_template = value
        self.OnPropertyChanged("SelectedTemplate")
        # notify VM to apply in Revit
        try:
            if self._on_template_selected and value is not None:
                self._on_template_selected(self, value)
        except:
            pass


class MainViewModel(ViewModel_BaseEventHandler):
    def __init__(self, external_event, handler):
        ViewModel_BaseEventHandler.__init__(self)
        self.external_event = external_event
        self.handler = handler

        # Data
        self._duplicated_views = ObservableCollection[ViewItem]()
        self._selected_items = []
        self._find_text = ""
        self._replace_text = ""
        # add
        self._prefix_text = ""
        self._suffix_text = ""

        # Commands
        self._dup_basic_cmd = DelegateCommand(lambda p=None: self._raise_duplicate("basic"))
        self._dup_detail_cmd = DelegateCommand(lambda p=None: self._raise_duplicate("detail"))
        self._dup_dependent_cmd = DelegateCommand(lambda p=None: self._raise_duplicate("dependent"))
        self._rename_cmd = DelegateCommand(self._raise_rename)
        # New: load selected views (no duplication)
        self._select_cmd = DelegateCommand(self._raise_load_selection)
        # New: remove selected rows from grid
        self._remove_cmd = DelegateCommand(self._remove_selected)

    # Collections
    @property
    def DuplicatedViews(self):
        return self._duplicated_views

    # Selection from grid
    @property
    def SelectedViewItems(self):
        return self._selected_items

    @SelectedViewItems.setter
    def SelectedViewItems(self, items):
        self._selected_items = list(items or [])
        self.OnPropertyChanged("SelectedViewItems")

    # Find/Replace
    @property
    def FindText(self):
        return self._find_text

    @FindText.setter
    def FindText(self, v):
        if self._find_text != (v or ""):
            self._find_text = v or ""
            self.OnPropertyChanged("FindText")

    @property
    def ReplaceText(self):
        return self._replace_text

    @ReplaceText.setter
    def ReplaceText(self, v):
        if self._replace_text != (v or ""):
            self._replace_text = v or ""
            self.OnPropertyChanged("ReplaceText")

    # Prefix/Suffix
    @property
    def PrefixText(self):
        return self._prefix_text

    @PrefixText.setter
    def PrefixText(self, v):
        if self._prefix_text != (v or ""):
            self._prefix_text = v or ""
            self.OnPropertyChanged("PrefixText")

    @property
    def SuffixText(self):
        return self._suffix_text

    @SuffixText.setter
    def SuffixText(self, v):
        if self._suffix_text != (v or ""):
            self._suffix_text = v or ""
            self.OnPropertyChanged("SuffixText")

    # Commands
    @property
    def DuplicateBasicCommand(self):
        return self._dup_basic_cmd

    @property
    def DuplicateDetailCommand(self):
        return self._dup_detail_cmd

    @property
    def DuplicateDependentCommand(self):
        return self._dup_dependent_cmd

    @property
    def RenameCommand(self):
        return self._rename_cmd

    # New: command property
    @property
    def SelectFromSelectionCommand(self):
        return self._select_cmd

    # New: Remove selected rows command
    @property
    def RemoveCommand(self):
        return self._remove_cmd

    # Called by handler to add rows after duplication
    def add_duplicated_view(self, element_id, name, template_name):
        item = ViewItem(element_id, name, template_name)
        # connect callback so UI selection triggers apply template
        item._on_template_selected = self._on_view_template_selected
        self._duplicated_views.Add(item)

    def set_view_templates(self, element_id, template_pairs):
        """Set the available templates (list of (ElementId, Name)) for the given view row."""
        try:
            # find row
            row = None
            for it in self._duplicated_views:
                if it.ElementId == element_id:
                    row = it
                    break
            if row is None:
                return
            # replace items
            row.Templates.Clear()
            for (tid, tname) in (template_pairs or []):
                row.Templates.Add(TemplateItem(tid, tname))
            # set current selection by name if possible
            current_name = row.ViewTemplate or ""
            sel = None
            for ti in row.Templates:
                if ti.Name == current_name:
                    sel = ti
                    break
            # if no match, try first item (e.g., "None")
            if sel is None and len(row.Templates) > 0:
                sel = row.Templates[0]
            row.SelectedTemplate = sel
        except Exception as ex:
            print("Error setting templates: {}".format(str(ex)))

    def update_view_template(self, element_id, template_name):
        """Update UI after applying a template in Revit."""
        try:
            for it in self._duplicated_views:
                if it.ElementId == element_id:
                    it.ViewTemplate = template_name or "none"
                    # also sync SelectedTemplate by name
                    for ti in it.Templates:
                        if ti.Name == it.ViewTemplate:
                            it.SelectedTemplate = ti
                            break
                    break
        except Exception as ex:
            print("Error updating view template: {}".format(str(ex)))

    # Internal: raise actions
    def _raise_duplicate(self, mode):
        try:
            self.handler.action = "duplicate"
            self.handler.duplicate_mode = mode  # "basic" | "detail" | "dependent"
            self.handler.view_model = self
            self.external_event.Raise()
        except Exception as ex:
            print("Error raising duplicate: {}".format(str(ex)))

    def _raise_rename(self, p=None):
        try:
            self.handler.action = "rename"
            self.handler.view_model = self
            self.external_event.Raise()
        except Exception as ex:
            print("Error raising rename: {}".format(str(ex)))

    # New: raise load selection action
    def _raise_load_selection(self, p=None):
        try:
            self.handler.action = "load_selection"
            self.handler.view_model = self
            self.external_event.Raise()
        except Exception as ex:
            print("Error raising load_selection: {}".format(str(ex)))

    # Callback from row when template changed in UI
    def _on_view_template_selected(self, view_item, template_item):
        try:
            # If multiple rows selected -> apply to all, else only this row
            selected = list(self._selected_items or [])
            if len(selected) > 1:
                self.handler.action = "apply_template_multi"
                self.handler.view_model = self
                self.handler.target_view_ids = [it.ElementId for it in selected]
            else:
                self.handler.action = "apply_template"
                self.handler.view_model = self
                self.handler.target_view_id = view_item.ElementId
            # pass template info
            self.handler.template_id = getattr(template_item, 'Id', None)
            self.handler.template_name = getattr(template_item, 'Name', None)
            self.external_event.Raise()
        except Exception as ex:
            print("Error raising apply_template: {}".format(str(ex)))

    # New: remove selected rows (no Revit change)
    def _remove_selected(self, p=None):
        try:
            to_remove = list(self._selected_items or [])
            if not to_remove:
                return
            # remove from collection
            for it in to_remove:
                try:
                    self._duplicated_views.Remove(it)
                except:
                    pass
            # clear selection after removal
            self.SelectedViewItems = []
        except Exception as ex:
            print("Error removing rows: {}".format(str(ex)))
