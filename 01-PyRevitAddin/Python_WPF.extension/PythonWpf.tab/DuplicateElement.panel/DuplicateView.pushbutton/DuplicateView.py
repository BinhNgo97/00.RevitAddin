import clr
clr.AddReference('ProtoGeometry')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
clr.AddReference("PresentationFramework")

from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.ObjectModel import ObservableCollection

clr.AddReferenceToFileAndPath("C:\\Program Files\\Autodesk\\Revit 2020\\AddIns\\DynamoForRevit\\IronPython.Wpf.dll")
import wpf

from System.Windows import Window, MessageBox

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument

# ---------------- Models for DataGrid ----------------
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

class ViewItem(INotifyPropertyChanged):
	def __init__(self, view_id, view_name, view_template_name):
		self._view_id = view_id
		self._view_name = view_name
		self._view_template = view_template_name
		self._handlers = []

	@property
	def ViewID(self): return self._view_id

	@property
	def ViewName(self): return self._view_name
	@ViewName.setter
	def ViewName(self, v):
		if self._view_name != v:
			self._view_name = v
			self._on_changed("ViewName")

	@property
	def ViewTemplate(self): return self._view_template
	@ViewTemplate.setter
	def ViewTemplate(self, v):
		if self._view_template != v:
			self._view_template = v
			self._on_changed("ViewTemplate")

	def add_PropertyChanged(self, h): self._handlers.append(h)
	def remove_PropertyChanged(self, h): self._handlers.remove(h)
	def _on_changed(self, name):
		for h in self._handlers:
			h(self, PropertyChangedEventArgs(name))

# ---------------- WPF Window ----------------
class MyWindow(Window):
	def __init__(self):
		# Load local XAML in this pushbutton folder
		self.win = wpf.LoadComponent(self, r"g:\My Drive\CV-STRUCTON\01-Dynamo\Python_WPF.extension\PythonWpf.tab\DuplicateElement.panel\DuplicateView.pushbutton\UI-DuplicateView.xaml")

		# Data
		self.items = ObservableCollection[ViewItem]()
		self.dgViews.ItemsSource = self.items

		# Wire buttons
		self.bt_Duplicate.Click += lambda s,e: self._duplicate(ViewDuplicateOption.Duplicate)
		self.bt_DuplicateWithDetail.Click += lambda s,e: self._duplicate(ViewDuplicateOption.WithDetailing)
		self.bt_DuplicateAsDependent.Click += lambda s,e: self._duplicate(ViewDuplicateOption.AsDependent)
		self.bt_Rename.Click += self._rename_selected
		self.bt_Cancel.Click += lambda s,e: self.Close()

	def _get_selected_views(self):
		ids = list(uidoc.Selection.GetElementIds())
		if not ids:
			MessageBox.Show("Please select one or more views in Revit, then try again.")
			return []
		views = []
		for id in ids:
			el = doc.GetElement(id)
			if isinstance(el, View) and not el.IsTemplate:
				views.append(el)
		if not views:
			MessageBox.Show("No valid views in the current selection.")
		return views

	def _vt_name(self, v):
		vtid = v.ViewTemplateId
		if vtid and vtid.IntegerValue > 0:
			# Might be a View or a ViewTemplate (View is fine)
			vt = doc.GetElement(vtid)
			return vt.Name if vt else "none"
		return "none"

	def _duplicate(self, option):
		# 1. collect current selection
		src_views = self._get_selected_views()
		if not src_views: return

		# 2. duplicate all in one transaction
		TransactionManager.Instance.EnsureInTransaction(doc)
		new_views = []
		for v in src_views:
			try:
				new_id = v.Duplicate(option)
				nv = doc.GetElement(new_id)
				new_views.append(nv)
			except Exception as ex:
				# Skip views that cannot be duplicated with this option
				pass
		TransactionManager.Instance.TransactionTaskDone()

		# 3. list in DataGrid
		for nv in new_views:
			self.items.Add(ViewItem(nv.Id, nv.Name, self._vt_name(nv)))

		if not new_views:
			MessageBox.Show("No views were duplicated with the chosen option.")

	def _rename_selected(self, sender, e):
		find_text = (self.tb_find.Text or "")
		replace_text = (self.tb_replace.Text or "")
		prefix_text = (self.tb_prefix.Text or "")
		suffix_text = (self.tb_suffix.Text or "")

		# Determine mode:
		# - If find_text provided => Find+Replace mode (replace_text can be empty)
		# - Else if prefix/suffix provided => Prefix/Suffix mode
		if not (find_text.strip() or prefix_text or suffix_text):
			MessageBox.Show("Enter Find Text, or Prefix/Suffix.")
			return

		selected = list(self.dgViews.SelectedItems)
		if not selected:
			MessageBox.Show("Select one or more rows in the table.")
			return

		TransactionManager.Instance.EnsureInTransaction(doc)
		renamed = 0
		failed = 0
		for item in selected:
			try:
				view = doc.GetElement(item.ViewID)
				if view is None:
					continue

				orig_name = view.Name or ""
				new_name = orig_name

				if find_text.strip():
					# Find+Replace mode (replace_text may be empty to delete)
					new_name = new_name.replace(find_text, replace_text)
				else:
					# Prefix/Suffix mode (apply to the whole name)
					if prefix_text:
						new_name = prefix_text + new_name
					if suffix_text:
						new_name = new_name + suffix_text

				if new_name and new_name != orig_name:
					view.Name = new_name
					item.ViewName = new_name
					renamed += 1
			except:
				failed += 1
				# Continue with next item

		TransactionManager.Instance.TransactionTaskDone()

		if renamed == 0 and failed == 0:
			MessageBox.Show("No view names were changed.")
		elif failed > 0:
			MessageBox.Show("Renamed: {}. Failed: {} (duplicate names or invalid characters).".format(renamed, failed))

# ---------------- Run ----------------
win = MyWindow()
win.ShowDialog()
# No OUT required; this tool works directly on the Revit model.
