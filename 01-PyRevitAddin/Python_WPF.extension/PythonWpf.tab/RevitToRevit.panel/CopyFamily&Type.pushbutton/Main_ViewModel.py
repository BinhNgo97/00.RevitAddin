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
    Parameter, ParameterType, StorageType, BuiltInParameter
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


class _NotUsed(object):
    pass


