import clr
clr.AddReference('ProtoGeometry')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
clr.AddReference("PresentationFramework")

import RevitServices
import Autodesk
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
view = doc.ActiveView
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument
app = uiapp.Application

from Autodesk.DesignScript.Geometry import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from System.Collections.Generic import *

clr.AddReferenceToFileAndPath("C:\\Program Files\\Autodesk\\Revit 2020\\AddIns\\DynamoForRevit\\IronPython.Wpf.dll")
import wpf 

from System.Windows import Window, Application, MessageBox, RoutedEventArgs
from System.Windows.Controls import Button, ListView, GridView, TextBox
from System.Windows.Forms import OpenFileDialog,DialogResult
from System.Windows.Controls import *
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.Specialized import NotifyCollectionChangedEventArgs
from collections import OrderedDict
from System import Action



class ViewItem(INotifyPropertyChanged):
	def __init__(self, view_id, view_name, view_type, view_template):
		self._view_id = view_id
		self._view_name = view_name
		self._view_type = view_type
		self._view_template = view_template
		self._propertyChangedHandlers = []

	@property
	def ViewID(self):
		return self._view_id

	@ViewID.setter
	def ViewID(self, value):
		if self._view_id != value:
			self._view_id = value
			self.OnPropertyChanged("ViewID")

	@property
	def ViewName(self):
		return self._view_name

	@ViewName.setter
	def ViewName(self, value):
		if self._view_name != value:
			self._view_name = value
			self.OnPropertyChanged("ViewName")

	@property
	def ViewType(self):
		return self._view_type

	@ViewType.setter
	def ViewType(self, value):
		if self._view_type != value:
			self._view_type = value
			self.OnPropertyChanged("ViewType")

	@property
	def ViewTemplate(self):
		return self._view_template

	@ViewTemplate.setter
	def ViewTemplate(self, value):
		if self._view_template != value:
			self._view_template = value
			self.OnPropertyChanged("ViewTemplate")

	def OnPropertyChanged(self, propertyName):
		for handler in self._propertyChangedHandlers:
			handler(self, PropertyChangedEventArgs(propertyName))

	def add_PropertyChanged(self, handler):
		self._propertyChangedHandlers.append(handler)

	def remove_PropertyChanged(self, handler):
		self._propertyChangedHandlers.remove(handler)

class MyWindow(Window):
	def __init__(self):
		self.winLoad = wpf.LoadComponent(self, r"G:\BINHNGO\Dynamo\000-Addin\BN_Dynamo\DynamoTools.extension\DynamoGeneral.tab\WorkwithView.panel\edit4.stack\DuplicateView.pushbutton\UI-DuplicateView.xaml")

		# ObservableCollection for DataGrid
		self.ViewItem = ObservableCollection[ViewItem]()
		self.ViewTemplates = []
		self.DataContext = self
		self.dataGrid_NewView.ItemsSource = self.ViewItem
		self.dataGrid_NewView.CurrentCellChanged += self.dataGrid_CurrentCellChanged

		self.cb_ViewType = self.winLoad.FindName("cb_ViewType")
		self.cb_ViewType.SelectionChanged += self.cb_selected
		self.cb_ViewType.ItemsSource = []

		self.tb_Find = self.winLoad.FindName("tb_find")
		self.tb_Replace = self.winLoad.FindName("tb_replace")


		self.bt_Cancel = self.winLoad.FindName("bt_Cancel")
		self.bt_Cancel.Click +=  self.cancel_click

		self.bt_DuplicateView = self.winLoad.FindName("bt_DuplicateView")
		self.bt_DuplicateView.Click +=  self.duplicate_click

		# Radio buttons
		self.rb_Duplicate = self.winLoad.FindName("rb_Duplicate")
		self.rb_DuplicateWithDetailing = self.winLoad.FindName("rb_DuplicateWithDetailing")
		self.rb_DuplicateAsDependent = self.winLoad.FindName("rb_DuplicateAsDependent")

		# ObservableCollection to automatically update ListView
		self.views_collection = ObservableCollection[str]()
		# self.lv_ViewbyViewType = self.winLoad.FindName("lv_ViewbyViewType")
		self.lv_ViewbyViewType.ItemsSource = self.views_collection

		# Get all ViewFamilyTypes and Views
		self.view_family_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
		self.all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
		# Logger("-----",self.all_views)
		for x in [view.Name for view in self.all_views if view.IsTemplate]:
			self.ViewTemplates.append(x)
		# Create a dictionary to store ViewFamilyType along with corresponding Views
		self.view_type_dict = self.create_view_type_dict()

		# Update ComboBox with ViewFamilyTypes
		self.update_combobox()
		# Tìm và thiết lập ItemsSource cho ComboBox trong DataGridTemplateColumn
		self.out_viewoldid = []
		self.out_viewnewname = []
		self.out_viewtemplate = []

	def create_view_type_dict(self):
		view_type_dict = {}
		for view_family_type in self.view_family_types:
			view_family_type_id = view_family_type.Id
			view_family_type_name = view_family_type.LookupParameter("Type Name").AsString()
			views_in_family_type = [view for view in self.all_views if view.GetTypeId() == view_family_type_id]
			if views_in_family_type:
				views_in_family_type.sort(key=lambda x: x.Name)
				view_type_dict[view_family_type_name] = views_in_family_type
		sorted_view_type_dict = OrderedDict(sorted(view_type_dict.items()))

		return sorted_view_type_dict
		# pass

	def update_combobox(self):
		self.cb_ViewType.ItemsSource = list(self.view_type_dict.keys())
		# pass

	def cb_selected(self, sender, e):
		selected_view_type = self.cb_ViewType.SelectedItem
		self.views_collection.Clear()  # Clear the existing items in the ObservableCollection

		if selected_view_type in self.view_type_dict:
			views = self.view_type_dict[selected_view_type]
			for view in views:
				self.views_collection.Add(view.Name)
		# pass

	def add_click(self, sender, e):
		selected_items = self.lv_ViewbyViewType.SelectedItems
		selected_view_type = self.cb_ViewType.SelectedItem
		search_text = self.tb_Find.Text
		replace_text = self.tb_Replace.Text

		for item in selected_items:
			# Find the view object from the name
			view = next((v for v in self.view_type_dict[selected_view_type] if v.Name == item), None)
			if not view:
				continue

			original_view_id = view.Id
			original_view_name = view.Name

			if search_text and replace_text:
				new_view_name = original_view_name.replace(search_text, replace_text)
			else:
				new_view_name = original_view_name

			new_view_item = ViewItem(original_view_id, new_view_name, selected_view_type, '')
			self.ViewItem.Add(new_view_item)

			self.out_viewoldid.append(original_view_id)
			self.out_viewnewname.append(new_view_name)

	def dataGrid_CurrentCellChanged(self, sender, e):
		dataGrid_NewView = sender
		selectedItems = list(dataGrid_NewView.SelectedItems)
		if dataGrid_NewView.CurrentCell.Column is not None:
			dataGrid_NewView.BeginEdit()

	def ComboBox_SelectionChanged(self, sender, e):
		dataGrid_NewView = self.dataGrid_NewView
		selectedItems = list(dataGrid_NewView.SelectedItems)
		selectedviewtemplate = sender.SelectedItem
		for x in selectedItems:
			x.ViewTemplate = selectedviewtemplate
			self.out_viewtemplate.append(x.ViewTemplate)

	def remove_click(self, sender, e):
		selected_items = list(self.dataGrid_NewView.SelectedItems)
		for item in selected_items:
			self.ViewItem.Remove(item)

	def cancel_click(self, sender, e):
		global OUT
		OUT = []
		self.Close()

	def duplicate_click(self, sender, e):
		global OUT

		duplicate_option = None
		if self.rb_Duplicate.IsChecked:
			duplicate_option = "Duplicate"
		elif self.rb_DuplicateWithDetailing.IsChecked:
			duplicate_option = "Duplicate with Detailing"
		elif self.rb_DuplicateAsDependent.IsChecked:
			duplicate_option = "Duplicate as a Dependent"
		OUT = {"View Old id": self.out_viewoldid,
				 "View New Name": self.out_viewnewname,
				  "View Template": self.out_viewtemplate,
				  "Duplicate Option": duplicate_option}
		self.Close()

myWindow = MyWindow()
myWindow.ShowDialog()
