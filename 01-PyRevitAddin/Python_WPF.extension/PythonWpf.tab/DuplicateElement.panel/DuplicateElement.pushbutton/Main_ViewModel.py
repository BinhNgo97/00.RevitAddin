# -*- coding: utf-8 -*-
# ===========================
# File: Main_ViewModel.py - Duplicate Element Type
# ===========================

import sys
import os
import clr

# Add parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

# Add Revit API references
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Data')

from aGeneral.ViewModel_Base import ViewModel_BaseEventHandler
from aGeneral.Command_Base import DelegateCommand
from System.Collections.ObjectModel import ObservableCollection
from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementType, FamilySymbol, 
    Parameter, StorageType, BuiltInParameter
)
from System import String
from System.Data import DataTable
from System.Collections.Generic import List as GenericList
from services_revit import (
    check_type_exists_in_category,
    get_families_by_category,
    get_element_types_by_category,
    get_first_symbol_of_family,
    get_parameter_value as revit_get_param_value,
    get_parameter_type as revit_get_param_type,
    compare_element_params_with_excel,
    update_element_type_parameters,
    get_element_type_by_name,
    categorize_types_from_excel,
    is_loadable_family_category
)
from services_excel import (
    load_excel_sheets as excel_list_sheets,
    read_sheet_rows,
    trim_data_rows,
    build_datatable_from_rows,
    extract_excel_param_dict
)


class TypeItem(ViewModel_BaseEventHandler):
    def __init__(self, name, element_type):
        ViewModel_BaseEventHandler.__init__(self)
        self._name = name
        self._element_type = element_type

    @property
    def Name(self):
        return self._name

    @property
    def ElementType(self):
        return self._element_type


class ParameterItem(ViewModel_BaseEventHandler):
    def __init__(self, name, value, parameter_type="Text"):
        ViewModel_BaseEventHandler.__init__(self)
        self._name = name
        self._value = value
        self._parameter_type = parameter_type

    @property
    def Name(self):
        return self._name

    @property
    def Value(self):
        return self._value

    @property
    def ParameterType(self):
        return self._parameter_type


class ExcelDataItem(ViewModel_BaseEventHandler):
    def __init__(self, row_data):
        ViewModel_BaseEventHandler.__init__(self)
        self._row_data = row_data

    @property
    def RowData(self):
        return self._row_data


class MainViewModel(ViewModel_BaseEventHandler):
    def __init__(self, external_event, handler):
        ViewModel_BaseEventHandler.__init__(self)
        self.external_event = external_event
        self.handler = handler
        
        # Collections
        self._categories = ObservableCollection[TypeItem]()
        self._fael_types = ObservableCollection[TypeItem]()  # unified Fa/El collection for single dropdown
        self._parameters = ObservableCollection[ParameterItem]()
        self._sheet_names = ObservableCollection[String]()
        self._excel_data_rows = ObservableCollection[ExcelDataItem]()
        self._data_table = None
        
        # Properties
        self._selected_category = None
        self._selected_type = None
        self._new_type_name = ""
        self._excel_file_path = ""
        self._selected_sheet = ""
        self._selected_fael = None
        
        # Commands
        self.select_excel_command = DelegateCommand(self.select_excel_file)
        self.import_data_command = DelegateCommand(self.import_data_to_grid)
        self.create_type_command = DelegateCommand(self.create_duplicate_type)
        # thêm Refresh command để khớp với script.py
        self.refresh_data_command = DelegateCommand(self.refresh_excel_data)
        # Command for updating element type parameters
        self.update_type_command = DelegateCommand(self.update_element_type)
        # Command for selecting elements from Revit
        self.select_row_elements_command = DelegateCommand(self.select_row_elements)
        # Thêm Command cho button chung
        self.action_command = DelegateCommand(self.perform_action)
        # New: bulk commands
        self.update_all_command = DelegateCommand(self.update_all_from_grid)
        self.create_all_new_command = DelegateCommand(self.create_all_new_from_grid)
        
        # Initialize
        self.load_categories()

        # Add missing initialization
        self._element_types = ObservableCollection[TypeItem]()  # For compatibility with load_types_by_family
        self._all_parameters = []  # For sorting parameters before adding to observable collection

        # Add properties to store categorized types
        self._categorized_types = None
        
    # Essential Properties for binding
    @property
    def Categories(self):
        return self._categories

    @property
    def FaElTypes(self):
        return self._fael_types

    @property
    def Parameters(self):
        return self._parameters

    @property
    def SheetNames(self):
        return self._sheet_names

    @property
    def DataTableDefaultView(self):
        try:
            return None if self._data_table is None else self._data_table.DefaultView
        except:
            return None

    @property
    def SelectedCategory(self):
        return self._selected_category

    @SelectedCategory.setter
    def SelectedCategory(self, value):
        if self._selected_category != value:
            self._selected_category = value
            self.OnPropertyChanged("SelectedCategory")
            self.load_types_by_category()

    @property
    def SelectedFaEl(self):
        return self._selected_fael

    @SelectedFaEl.setter
    def SelectedFaEl(self, value):
        if self._selected_fael != value:
            self._selected_fael = value
            self.OnPropertyChanged("SelectedFaEl")
            self.load_type_parameters_from_fael()

    @property
    def SelectedType(self):
        return self._selected_type

    @SelectedType.setter
    def SelectedType(self, value):
        if self._selected_type != value:
            self._selected_type = value
            self.OnPropertyChanged("SelectedType")

    @property
    def NewTypeName(self):
        return self._new_type_name

    @NewTypeName.setter
    def NewTypeName(self, value):
        if self._new_type_name != value:
            self._new_type_name = value
            self.OnPropertyChanged("NewTypeName")

    @property
    def ExcelFilePath(self):
        return self._excel_file_path

    @ExcelFilePath.setter
    def ExcelFilePath(self, value):
        if self._excel_file_path != value:
            self._excel_file_path = value
            self.OnPropertyChanged("ExcelFilePath")

    @property
    def SelectedSheet(self):
        return self._selected_sheet

    @SelectedSheet.setter
    def SelectedSheet(self, value):
        if self._selected_sheet != value:
            self._selected_sheet = value
            self.OnPropertyChanged("SelectedSheet")

    # Commands
    @property
    def SelectExcelCommand(self):
        return self.select_excel_command

    @property
    def ImportDataCommand(self):
        return self.import_data_command

    @property
    def CreateTypeCommand(self):
        return self.create_type_command

    # thêm property cho RefreshDataCommand
    @property
    def RefreshDataCommand(self):
        return self.refresh_data_command

    # Add property for UpdateTypeCommand
    @property
    def UpdateTypeCommand(self):
        return self.update_type_command

    # Command for selecting elements from Revit
    @property
    def SelectRowElementsCommand(self):
        return self.select_row_elements_command

    # Thêm property cho ActionCommand
    @property
    def ActionCommand(self):
        return self.action_command

    # New command properties
    @property
    def UpdateAllCommand(self):
        return self.update_all_command

    @property
    def CreateAllNewCommand(self):
        return self.create_all_new_command
    
    def load_categories(self):
        """
        Loads and filters categories from the active Revit document.
        This method retrieves all categories from the current Revit document,
        excluding those with specific file suffixes (e.g., .dwg, .rvt, .ifc, .dxf, .sat, .skp)
        and certain unwanted category names (e.g., Analytical, Areas, Assemblies, Project Information).
        Only categories that allow bound parameters are included, unless the check fails,
        in which case the category is added anyway. The filtered categories are sorted
        alphabetically and added to the observable collection for use in the UI.
        Prints the number of loaded categories or an error message if loading fails.
        """
        """Load all categories from Revit document, filtering out unwanted ones like .dwg, .rvt"""

        try:
            doc = __revit__.ActiveUIDocument.Document
            if not doc:
                return
            self._categories.Clear()
            categories = list(doc.Settings.Categories)
            cat_items = []
            # Filter categories
            excluded_suffixes = [".dwg", ".rvt", ".ifc", ".dxf", ".sat", ".skp"]
            excluded_names = ["Analytical", "Areas", "Assemblies", "Project Information"]
            for cat in categories:
                try:
                    if cat and cat.Name:
                        # Skip categories with excluded suffixes or names
                        if any(cat.Name.endswith(suffix) for suffix in excluded_suffixes):
                            continue
                        if any(excluded in cat.Name for excluded in excluded_names):
                            continue
                        # Only include categories that can have types
                        try:
                            if cat.AllowsBoundParameters:
                                cat_items.append(TypeItem(cat.Name, cat))
                        except:
                            # If AllowsBoundParameters check fails, add anyway
                            cat_items.append(TypeItem(cat.Name, cat))
                except:
                    pass
            # Sort categories alphabetically by name
            cat_items.sort(key=lambda x: x.Name)
            # Add to observable collection
            for item in cat_items:
                self._categories.Add(item)
            # print("Loaded {} filtered categories".format(len(cat_items)))
        except Exception as ex:
            # print("Error loading categories: {}".format(str(ex)))
            pass

    def load_types_by_category(self):
        """Load appropriate types or families for the selected category into dropdown"""
        try:
            doc = __revit__.ActiveUIDocument.Document
            if not doc or not self._selected_category:
                return
            self._fael_types.Clear()
            self._parameters.Clear()
            self._selected_fael = None
            self.OnPropertyChanged("SelectedFaEl")
            cat = self._selected_category.ElementType
            if not cat:
                return
            # Get appropriate elements (families or types) based on category type
            # The function handles both system and loadable family categories internally
            name_element_tuples = get_element_types_by_category(doc, cat, return_tuple=True)
            # Add items to dropdown
            for name, element in name_element_tuples:
                self._fael_types.Add(TypeItem(name, element))
            # Auto-select first item if any
            if self._fael_types.Count > 0:
                self.SelectedFaEl = self._fael_types[0]
                # print("Selected first item: {}".format(self._fael_types[0].Name))
            else:
                # print("No items found for category: {}".format(getattr(cat, "Name", "")))
                pass
        except Exception as ex:
            # print("Error loading types: {}".format(str(ex)))
            import traceback
            # print(traceback.format_exc())
            return
        
    def load_type_parameters_from_fael(self):
        """Load parameters from selected family or element type"""
        try:
            self._parameters.Clear()
            if not self._selected_fael:
                return

            doc = __revit__.ActiveUIDocument.Document
            elem = self._selected_fael.ElementType
            
            if not elem:
                # print("No element found in selected item")
                return
                
            # Get element name safely for logging
            elem_name = "Unknown"
            try:
                if hasattr(elem, "Name"):
                    elem_name = elem.Name
                elif hasattr(elem, "get_Parameter"):
                    name_param = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if name_param:
                        elem_name = name_param.AsString()
            except:
                pass
                
            # Handle families differently from direct element types
            if hasattr(elem, "GetFamilySymbolIds"):  # It's a Family
                # Get the first symbol (type) from this family
                sym = get_first_symbol_of_family(doc, elem)
                if sym:
                    # Store the element type for later use
                    self.SelectedType = TypeItem(sym.Name if hasattr(sym, "Name") else "Unknown", sym)
                    
                    # Load parameters from the symbol
                    self._load_parameters_from_element(sym)
                    # print("Loaded parameters from family symbol: {}".format(
                    #     sym.Name if hasattr(sym, "Name") else "Unknown"))
                else:
                    # print("No symbol found for family: {}".format(elem_name))
                    pass
            else:  # It's already an ElementType
                # Store the element type directly
                self.SelectedType = self._selected_fael
                
                # Load parameters from the element type
                self._load_parameters_from_element(elem)
                # print("Loaded parameters from element type: {}".format(elem_name))
        except Exception as ex:
            # print("Error loading parameters: {}".format(str(ex)))
            import traceback
            # print(traceback.format_exc())

    # Wrapper helpers to use services_revit functions
    def get_parameter_value(self, param):
        try:
            return revit_get_param_value(param)
        except:
            return ""

    def get_parameter_type(self, param):
        try:
            return revit_get_param_type(param)
        except:
            return "Unknown"

    def _load_parameters_from_element(self, elem):
        """Helper to load parameters from an element into the Parameters collection"""
        try:
            if not elem:
                return
                
            param_items = []
            for p in elem.Parameters:
                if p and p.Definition and p.Definition.Name:
                    val = self.get_parameter_value(p)
                    ptype = self.get_parameter_type(p)
                    param_items.append(ParameterItem(p.Definition.Name, val, ptype))

            # Sort parameters alphabetically
            param_items.sort(key=lambda x: (x.Name or "").lower())
            
            # Add to observable collection
            for it in param_items:
                self._parameters.Add(it)
            # print("Loaded {} parameters".format(len(param_items)))
        except Exception as ex:
            # print("Error in _load_parameters_from_element: {}".format(str(ex)))
            pass

    def check_type_exists_in_revit(self, type_name):
        """Check if a type name already exists in Revit (by selected category)"""
        try:
            if not type_name or not self._selected_category:
                return False
            doc = __revit__.ActiveUIDocument.Document
            cat = self._selected_category.ElementType
            return check_type_exists_in_category(doc, cat, type_name)
        except:
            return False

    def select_excel_file(self, parameter=None):
        """Open file dialog to select Excel file"""
        try:
            from pyrevit import forms
            
            # Use pyrevit forms for file selection
            excel_file = forms.pick_file(
                title="Select Excel File",
                file_ext="xlsx;xls",
                restore_dir=True
            )
            
            if excel_file:
                self.ExcelFilePath = excel_file
                # clear previous state
                self._excel_data_rows.Clear()
                self._data_table = None
                self.OnPropertyChanged("DataTableDefaultView")
                self.load_excel_sheets()
                
        except Exception as ex:
            # print("Error selecting Excel file: {}".format(str(ex)))
            pass

    def load_excel_sheets(self):
        """Load available sheets from Excel file"""
        try:
            if not self._excel_file_path:
                # print("No Excel file path set")
                return

            self._sheet_names.Clear()
            # print("Loading sheets from: {}".format(self._excel_file_path))
            sheets = excel_list_sheets(self._excel_file_path)

            if sheets and len(sheets) > 0:
                for sheet in sheets:
                    self._sheet_names.Add(sheet)
                # Auto-select first sheet
                self.SelectedSheet = self._sheet_names[0]
                # print("Loaded {} sheets".format(self._sheet_names.Count))
            else:
                # print("No sheets detected. Please check the file.")
                pass
        except Exception as ex:
            # print("Error loading Excel sheets: {}".format(str(ex)))
            pass

    def import_data_to_grid(self, parameter=None):
        """Import selected sheet data to DataGrid using services_excel (simple path)."""
        try:
            # print("Import data to grid called")
            if not self._excel_file_path:
                # print("No Excel file selected")
                return
            if not self._selected_sheet:
                # auto pick first if available
                if self._sheet_names.Count > 0:
                    self.SelectedSheet = self._sheet_names[0]
                else:
                    # print("No sheet selected")
                    return

            # Read rows directly from file
            rows = read_sheet_rows(self._excel_file_path, self._selected_sheet)
            if not rows:
                # print("No data rows read from sheet '{}'".format(self._selected_sheet))
                return

            # Trim rows to non-empty rectangular region
            rows = trim_data_rows(rows)
            if not rows:
                # print("No usable data after trim")
                return

            # Determine Name/Status columns
            name_col_index = None
            status_col_index = None
            header_row = rows[0] if rows else []
            for i, col_name in enumerate(header_row):
                try:
                    s = str(col_name).strip().lower() if col_name is not None else ""
                    if s == "name":
                        name_col_index = i
                    elif s == "status":
                        status_col_index = i
                except:
                    pass
            if name_col_index is None:
                # print("WARNING: Could not find Name column. Using column 0.")
                name_col_index = 0

            # Revit doc & category
            doc = __revit__.ActiveUIDocument.Document
            cat = self._selected_category.ElementType if self._selected_category else None
            if not doc or not cat:
                # print("No document or category selected")
                return

            # Exists/check callbacks
            def _exists(name):
                return check_type_exists_in_category(doc, cat, name)

            def _compare_params(type_name, header_row, data_row, col_idx=name_col_index):
                return compare_element_params_with_excel(doc, cat, type_name, header_row, data_row, col_idx)

            # Build DataTable
            dt = build_datatable_from_rows(
                rows,
                add_status=(status_col_index is None),
                check_exists_fn=_exists,
                param_comparison_fn=_compare_params,
                name_col_index=name_col_index,
                status_col_index=status_col_index
            )

            self._data_table = dt
            self.OnPropertyChanged("DataTableDefaultView")
            # print("Imported {} rows, {} columns".format(dt.Rows.Count, dt.Columns.Count))
        except Exception as ex:
            # print("Error importing data: {}".format(str(ex)))
            import traceback
            # print(traceback.format_exc())

    def refresh_excel_data(self, parameter=None):
        try:
            if self._selected_sheet:
                self.import_data_to_grid()
            else:
                # print("No sheet selected")
                pass
        except Exception as ex:
            # print("Error refreshing Excel data: {}".format(str(ex)))
            pass

    def update_element_type(self, parameter=None):
        """Update element type parameters from Excel data"""
        try:
            # Get selected row and status
            if not parameter or not hasattr(parameter, "Row"):
                # print("No row selected for update")
                return
                 
            row = parameter.Row
            type_name = str(row["Name"])
            status = self._get_row_status(row)
            if status != "Update":
                # print("Selected type doesn't need updating")
                return
                 
            # Get category and document
            doc = __revit__.ActiveUIDocument.Document
            cat = self._selected_category.ElementType if self._selected_category else None
            if not doc or not cat:
                # print("No document or category selected")
                return
                 
            # Get element type
            element_type = get_element_type_by_name(doc, cat, type_name)
            if not element_type:
                # print("Element type not found: {}".format(type_name))
                return
                 
            # Build parameter dictionary from row (update all provided params)
            param_dict = {}
            for i in range(self._data_table.Columns.Count):
                col_name = self._data_table.Columns[i].ColumnName
                if col_name not in ["Name", "Status", "Action"]:
                    param_dict[col_name] = row[i]
        
            # Update parameters
            self.handler.element_type = element_type
            self.handler.param_dict = param_dict
            self.handler.action = "update_element_parameters"
            self.external_event.Raise()
            # print("Element type '{}' update requested".format(type_name))
            # Optimistically update status in grid
            self._set_row_status(row, "Existing")
        except Exception as ex:
            # print("Error updating element type: {}".format(str(ex)))
            pass

    def select_row_elements(self, parameter=None):
        """No-op: selection disabled; kept to satisfy existing command wiring."""
        try:
            # print("SelectRowElements is disabled. No action performed.")
            pass
        except:
            pass

    def perform_action(self, parameter):
        """Perform action based on row status"""
        try:
            if not parameter or not hasattr(parameter, "Row"):
                # print("No row selected for action")
                return
                
            row = parameter.Row
            status = self._get_row_status(row)

            if status == "Update":
                self.update_element_type(parameter)
            elif status == "New":
                # Create a new type from the currently selected source type and row values
                if not self._selected_type:
                    # print("Please select a source type to duplicate from")
                    return
                # Reuse existing duplication flow with row-based parameters
                self.create_duplicate_type(parameter)
            elif status == "Existing":
                # No operation
                # print("No action for 'Existing' status")
                pass
            else:
                # print("Unknown status: {}".format(status))
                pass
        except Exception as ex:
            # print("Error performing action: {}".format(str(ex)))
            pass

    def create_duplicate_type(self, parameter=None):
        """Create duplicate element type from selected row or create all new types"""
        try:
            # If parameter is a DataRowView => create from row values
            if parameter is not None and hasattr(parameter, "Row"):
                row = parameter.Row
                new_type_name = str(row["Name"])
                if not self._selected_type:
                    # print("Please select a source type to duplicate from")
                    return
                self.handler.source_type = self._selected_type.ElementType
                self.handler.new_type_name = new_type_name
                self.handler.parameter_dict = self._extract_parameters_from_row(row)
                self.handler.action = "create_duplicate_type"
                self.external_event.Raise()
                # print("Create type requested: {}".format(new_type_name))
                # Update status in grid
                self._set_row_status(row, "Existing")
                return

            # Fallback: original manual flow using SelectedType and NewTypeName
            if not self._selected_type or not self._new_type_name:
                # print("Please select a type and enter a new name")
                return
            self.handler.source_type = self._selected_type.ElementType
            self.handler.new_type_name = self._new_type_name
            self.handler.excel_data_rows = self._excel_data_rows
            self.handler.action = "create_duplicate_type"
            self.external_event.Raise()
            
        except Exception as ex:
            # print("Error creating duplicate type: {}".format(str(ex)))
            pass

    def _extract_parameters_from_row(self, row):
        """Extract parameters from a DataRow"""
        param_dict = {}
        try:
            for i in range(self._data_table.Columns.Count):
                col_name = self._data_table.Columns[i].ColumnName
                if col_name not in ["Name", "Status", "Action"]:
                    param_dict[col_name] = row[i]
        except Exception as ex:
            # print("Error extracting parameters: {}".format(str(ex)))
            pass
        return param_dict

    def update_all_from_grid(self, parameter=None):
        """Update all rows that have Status == 'Update'"""
        try:
            if self._data_table is None or self._data_table.Rows.Count == 0:
                # print("No data to update")
                return
            doc = __revit__.ActiveUIDocument.Document
            cat = self._selected_category.ElementType if self._selected_category else None
            if not doc or not cat:
                # print("No document or category selected")
                return

            # Iterate rows with Status == Update (use Action as fallback)
            for i in range(self._data_table.Rows.Count):
                row = self._data_table.Rows[i]
                status = self._get_row_status(row)
                if status != "Update":
                    continue

                type_name = str(row["Name"]) if "Name" in [c.ColumnName for c in self._data_table.Columns] else ""
                if not type_name:
                    continue
                element_type = get_element_type_by_name(doc, cat, type_name)
                if not element_type:
                    # print("Element type not found for update: {}".format(type_name))
                    continue

                # Build param dict for row
                param_dict = {}
                for c in range(self._data_table.Columns.Count):
                    col = self._data_table.Columns[c].ColumnName
                    if col not in ["Name", "Status", "Action"]:
                        param_dict[col] = row[c]

                # Raise external event to update
                self.handler.element_type = element_type
                self.handler.param_dict = param_dict
                self.handler.action = "update_element_parameters"
                self.external_event.Raise()
                # print("Queued update for '{}'".format(type_name))

                # Optimistically set status to Existing
                self._set_row_status(row, "Existing")
        except Exception as ex:
            # print("Error in update_all_from_grid: {}".format(str(ex)))
            pass

    def create_all_new_from_grid(self, parameter=None):
        """Create all new types for rows with Status == 'New' using batch handler"""
        try:
            if self._data_table is None or self._data_table.Rows.Count == 0:
                # print("No data to create")
                return
            if not self._selected_type:
                # print("Please select a source type to duplicate from")
                return

            # Collect new rows
            new_items = []
            for i in range(self._data_table.Rows.Count):
                row = self._data_table.Rows[i]
                status = self._get_row_status(row)
                if status != "New":
                    continue
                type_name = str(row["Name"]) if "Name" in [c.ColumnName for c in self._data_table.Columns] else ""
                if not type_name:
                    continue
                new_items.append({'name': type_name, 'row': [row[j] for j in range(self._data_table.Columns.Count)]})

            if not new_items:
                # print("No 'New' rows to create")
                return

            # Prepare handler for batch create
            self.handler.source_type = self._selected_type.ElementType
            header_row = [self._data_table.Columns[c].ColumnName for c in range(self._data_table.Columns.Count)]
            self.handler.header_row = header_row
            self.handler.new_types_info = new_items
            self.handler.action = "create_multiple_types"
            self.external_event.Raise()
            # print("Queued creation for {} new types".format(len(new_items)))
            # Optimistically set statuses to Existing
            for i in range(self._data_table.Rows.Count):
                row = self._data_table.Rows[i]
                if self._get_row_status(row) == "New":
                    self._set_row_status(row, "Existing")
        except Exception as ex:
            # print("Error in create_all_new_from_grid: {}".format(str(ex)))
            pass

    # Helpers to read/write status from either 'Status' or 'Action' column
    def _get_row_status(self, row):
        try:
            v = row["Status"]
            return "" if v is None else str(v).strip()
        except:
            try:
                v = row["Action"]
                return "" if v is None else str(v).strip()
            except:
                return ""

    def _set_row_status(self, row, value):
        try:
            row["Status"] = value
            return
        except:
            pass
        try:
            row["Action"] = value
        except:
            pass

    def get_status_counts(self):
        """Return counts of (existing, update, new) based on current DataTable."""
        try:
            if self._data_table is None or self._data_table.Rows.Count == 0:
                return (0, 0, 0)
            existing = 0
            update = 0
            new = 0
            for i in range(self._data_table.Rows.Count):
                row = self._data_table.Rows[i]
                s = self._get_row_status(row)
                if s == "Existing":
                    existing += 1
                elif s == "Update":
                    update += 1
                elif s == "New":
                    new += 1
            return (existing, update, new)
        except:
            return (0, 0, 0)


