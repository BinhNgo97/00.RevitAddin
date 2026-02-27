# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, Family, ElementType, StorageType, ElementId,
    BuiltInCategory, BuiltInParameter, Transaction, TransactionStatus,
    ElementCategoryFilter
)
from System import Enum
from contextlib import contextmanager

@contextmanager
def revit_transaction(doc, name="Transaction"):
    """Context manager for Revit transactions"""
    t = Transaction(doc, name)
    t.Start()
    try:
        yield t
        if t.GetStatus() == TransactionStatus.Started:
            t.Commit()
    except Exception:
        if t.GetStatus() == TransactionStatus.Started:
            t.RollBack()
        raise

def _to_float_try(value):
    """Try to convert a value to float, return None if not possible"""
    try:
        s = str(value).strip()
        if "," in s and "." not in s:
            s = s.replace(",", ".")
        return float(s)
    except:
        return None

def is_loadable_family_category(doc, category):
    """Determine if a category uses loadable families vs. system families"""
    if not doc or not category:
        return False

    # 1) If any Family belongs to this category -> loadable
    try:
        fams = FilteredElementCollector(doc).OfClass(Family).ToElements()
        for f in fams:
            try:
                if f.FamilyCategory and f.FamilyCategory.Id == category.Id:
                    return True
            except:
                pass
    except:
        pass

    # 2) Otherwise, try to get element types for the category
    types = []
    # First, try using the BuiltInCategory enum conversion
    try:
        bic = Enum.ToObject(BuiltInCategory, category.Id.IntegerValue)
        types = FilteredElementCollector(doc) \
            .OfCategory(bic) \
            .WhereElementIsElementType() \
            .ToElements()
    except:
        types = []

    # Fallback: collect all element types and filter by Category.Id
    if not types:
        try:
            all_types = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
            types = [t for t in all_types
                     if hasattr(t, "Category") and t.Category is not None and t.Category.Id == category.Id]
        except:
            types = []

    # 3) If we found types, check for Family property to decide
    for t in types:
        if hasattr(t, "Family"):
            return getattr(t, "Family", None) is not None
        else:
            # No Family property => system category
            return False

    # 4) Nothing determinative found => assume system family
    return False

def get_element_types_by_category(doc, category, return_tuple=True):
    """
    Get element types of a specific category, handling both loadable and system families consistently.
    Always returns element types (not families).
    
    Args:
        doc: Revit document
        category: Revit category
        return_tuple: If True, returns list of (name, element) tuples, otherwise returns just elements
        
    Returns:
        List of (name, element_type) tuples or just element_types
    """
    results = []
    if not doc or not category:
        return results
    
    try:
        # Determine if this is a loadable family category
        is_loadable = is_loadable_family_category(doc, category)
        built_in_cat = category.Id.IntegerValue

        from Autodesk.Revit.DB import BuiltInParameter

        if is_loadable:
            # For loadable family categories, get types through families
            # print("Processing loadable family category: {}".format(category.Name))
            families = get_families_by_category(doc, category)
            
            if families:
                for f in families:
                    if f and hasattr(f, "GetFamilySymbolIds"):
                        family_name = "Unknown Family"
                        try:
                            family_name = f.Name if hasattr(f, "Name") else "Unknown Family"
                        except:
                            pass
                        
                        # Get all types (symbols) from this family
                        symbol_ids = []
                        try:
                            symbol_ids = list(f.GetFamilySymbolIds())
                        except Exception as ex:
                            # print("Error getting symbol IDs: {}".format(str(ex)))
                            pass
                            
                        for symbol_id in symbol_ids:
                            try:
                                symbol = doc.GetElement(symbol_id)
                                if not symbol:
                                    continue
                                    
                                # Get element name
                                element_name = None
                                try:
                                    name_param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                                    if name_param:
                                        element_name = name_param.AsString()
                                except:
                                    pass
                                
                                if not element_name and hasattr(symbol, "Name"):
                                    try:
                                        element_name = symbol.Name
                                    except:
                                        pass
                                
                                if not element_name:
                                    element_name = "Unknown Type"
                                    
                                if return_tuple:
                                    results.append((element_name, symbol))
                                else:
                                    results.append(symbol)
                                # print("  - Added element type: {} ({})".format(element_name, family_name))
                            except Exception as ex:
                                # print("Error processing symbol: {}".format(str(ex)))
                                pass
                
                # print("Successfully added {} element types for loadable family category {}".format(len(results), category.Name))
                    
        else:
            # For system family categories, return element types directly
            # print("Processing system family category: {}".format(category.Name))
            try:
                types = []
                try:
                    # Use enum conversion properly
                    bic = Enum.ToObject(BuiltInCategory, built_in_cat)
                    types = FilteredElementCollector(doc) \
                        .OfCategory(bic) \
                        .WhereElementIsElementType() \
                        .ToElements()
                except:
                    # Fallback: manual filter by Category.Id
                    all_types = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
                    types = [t for t in all_types
                             if hasattr(t, "Category") and t.Category is not None and t.Category.Id == category.Id]

                for t in types:
                    try:
                        element_name = "Unknown Type"
                        try:
                            name_param = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                            if name_param:
                                element_name = name_param.AsString()
                        except:
                            pass
                        if (not element_name) and hasattr(t, "Name"):
                            try:
                                element_name = t.Name
                            except:
                                pass
                        if return_tuple:
                            results.append((element_name, t))
                        else:
                            results.append(t)
                        # print("  - Added element type: {}".format(element_name))
                    except Exception as ex:
                        # print("Error processing type: {}".format(str(ex)))
                        pass
                # print("Successfully added {} element types for category {}".format(len(results), category.Name))
            except Exception as ex:
                # print("Error collecting element types: {}".format(str(ex)))
                pass
            
        # Sort results by name
        if return_tuple:
            results.sort(key=lambda x: x[0].lower() if x[0] else "")
        else:
            results.sort(key=lambda x: x.Name.lower() if hasattr(x, "Name") and x.Name else "")
            
    except Exception as ex:
        # print("Error in get_element_types_by_category: {}".format(str(ex)))
        import traceback
        # print(traceback.format_exc())
        pass
        
    return results

def get_all_family_types_in_category(doc, category):
    """
    Get all element types in a category, handling both direct types and family-based types
    Returns a list of tuples (name, element_type)
    
    This is a wrapper around get_element_types_by_category with return_tuple=True
    """
    # Directly call get_element_types_by_category with return_tuple=True
    return get_element_types_by_category(doc, category, return_tuple=True)

def check_type_exists_in_category(doc, category, type_name):
    """
    Check if a type name already exists in a category
    Handle both direct types and family-based types
    """
    try:
        if not doc or not category or not type_name:
            return False
        
        # Get all types in category
        all_types = get_all_family_types_in_category(doc, category)
        
        # Check if type_name exists
        for name, _ in all_types:
            # print("Checking type name:", name)
            if name == type_name:
                return True
        
        return False
    except:
        return False

def get_families_by_category(doc, category):
    items = []
    if not doc or not category:
        return items
    fams = FilteredElementCollector(doc).OfClass(Family).ToElements()
    for f in fams:
        try:
            if f and f.FamilyCategory and f.FamilyCategory.Id == category.Id:
                items.append(f)
        except:
            pass
    return items

def get_first_symbol_of_family(doc, family):
    if not doc or not family or not hasattr(family, "GetFamilySymbolIds"):
        return None
    try:
        for sid in family.GetFamilySymbolIds():
            sym = doc.GetElement(sid)
            if sym:
                return sym
    except:
        pass
    return None

def get_parameter_value(param, doc=None):
    try:
        if param is None:
            return ""
        try:
            val = param.AsValueString()
            if val not in (None, ""):
                return val
        except:
            pass

        st = getattr(param, "StorageType", None)
        if st == StorageType.String:
            return param.AsString() or ""
        if st == StorageType.Integer:
            return str(param.AsInteger())
        if st == StorageType.Double:
            return str(param.AsDouble())
        if st == StorageType.ElementId:
            try:
                eid = param.AsElementId()
                if eid is None:
                    return ""
                if doc is not None:
                    try:
                        el = doc.GetElement(eid)
                        if el is not None and hasattr(el, "Name"):
                            return el.Name or str(eid.IntegerValue)
                    except:
                        pass
                return str(eid.IntegerValue)
            except:
                return ""
        return param.AsValueString() or ""
    except:
        return ""

def get_parameter_type(param):
    try:
        if param is None:
            return "Unknown"
        try:
            ptype = param.Definition.ParameterType
        except:
            ptype = None

        # Best-effort map
        from Autodesk.Revit.DB import ParameterType as PT
        if ptype == PT.Text: return "Text"
        if ptype == PT.Number: return "Number"
        if ptype == PT.Length: return "Length"
        if ptype == PT.Area: return "Area"
        if ptype == PT.Volume: return "Volume"

        st = getattr(param, "StorageType", None)
        if st == StorageType.String: return "Text"
        if st == StorageType.Integer: return "Integer"
        if st == StorageType.Double: return "Double"
        if st == StorageType.ElementId: return "ElementId"
        return "Other"
    except:
        return "Unknown"

def get_element_type_by_name(doc, category, type_name):
    """
    Get element type by name in category
    Handle both direct types and family-based types
    """
    try:
        if not doc or not category or not type_name:
            return None
        
        # Get all types in category
        all_types = get_all_family_types_in_category(doc, category)
        
        # Find type by name
        for name, element_type in all_types:
            if name == type_name:
                return element_type
        
        return None
    except:
        return None

def get_element_parameter_dict(element):
    """Get dictionary of parameter name to value for an element"""
    param_dict = {}
    if not element:
        return param_dict
        
    try:
        for param in element.Parameters:
            if param and param.Definition and param.Definition.Name:
                param_dict[param.Definition.Name] = get_parameter_value(param)
    except:
        pass
        
    return param_dict

def compare_element_params_with_excel(doc, category, type_name, excel_headers, excel_row, name_col_index=None):
    """
    Compare element parameters with Excel data
    Returns list of different parameter names
    
    Args:
        doc: Revit document
        category: Revit category
        type_name: Name of the type to compare
        excel_headers: List of column headers from Excel
        excel_row: List of values from one Excel row
        name_col_index: Index of the column containing type names (default: 0)
    """
    if not doc or not category or not type_name or not excel_headers or not excel_row:
        return []
        
    try:
        # Determine the index of the Name column if not provided
        if name_col_index is None:
            # Try to find a column named "Name"
            for i, col_name in enumerate(excel_headers):
                if col_name and str(col_name).strip().lower() == "name":
                    name_col_index = i
                    break
            
            # If still not found, default to the first column
            if name_col_index is None:
                name_col_index = 0
        
        # Lấy element type trong Revit dựa vào tên (type_name)
        element_type = get_element_type_by_name(doc, category, type_name)
        if not element_type:
            # print("Type not found: {}".format(type_name))
            return []
            
        # Lấy tất cả các parameters của element type trong Revit
        element_params = get_element_parameter_dict(element_type)
        
        # Extract parameters từ một hàng Excel (bỏ qua cột Name)
        excel_params = {}
        for i in range(len(excel_headers)):
            # Skip the Name column
            if i == name_col_index:
                continue
            if i < len(excel_headers) and i < len(excel_row) and excel_headers[i] and str(excel_headers[i]).strip():
                header_name = str(excel_headers[i]).strip()
                # Skip Status and Action columns
                if header_name.lower() in ["status", "action"]:
                    continue
                # Include Excel values even if empty string; we still want to compare against Revit
                excel_value = "" if excel_row[i] is None else str(excel_row[i]).strip()
                excel_params[header_name] = excel_value
        
        # So sánh parameters và thu thập các parameters khác nhau
        different_params = []
        for param_name, excel_value in excel_params.items():
            if param_name in element_params:
                element_value = element_params[param_name]
                # Try numeric tolerant compare first
                f1 = _to_float_try(element_value)
                f2 = _to_float_try(excel_value)
                if (f1 is not None) and (f2 is not None):
                    if abs(f1 - f2) > 1e-6:
                        different_params.append(param_name)
                        # print("Parameter '{}' differs (numeric): Revit='{}', Excel='{}'".format(
                        #     param_name, element_value, excel_value))
                else:
                    if str(element_value).strip() != str(excel_value).strip():
                        different_params.append(param_name)
                        # print("Parameter '{}' differs: Revit='{}', Excel='{}'".format(
                        #     param_name, element_value, excel_value))
            else:
                # print("Parameter '{}' not found in element type".format(param_name))
                pass
        return different_params
    except Exception as ex:
        # print("Error comparing parameters: {}".format(str(ex)))
        import traceback
        # print(traceback.format_exc())
        return []

def update_element_type_parameters(doc, element_type, param_dict):
    """Update element type parameters with values from dict"""
    if not doc or not element_type or not param_dict:
        return False

    def _to_bool_int(v):
        try:
            s = str(v).strip().lower()
            if s in ("1", "true", "yes", "y", "on"):
                return 1
            if s in ("0", "false", "no", "n", "off"):
                return 0
            return 1 if float(s) != 0.0 else 0
        except:
            return 0

    def _to_float(v):
        try:
            s = str(v).strip()
            if "," in s and "." not in s:
                s = s.replace(",", ".")
            return float(s)
        except:
            return None

    success = False
    try:
        # Start transaction
        with revit_transaction(doc, "Update Type Parameters"):
            for param in element_type.Parameters:
                try:
                    if not param or not param.Definition or not param.Definition.Name:
                        continue
                    if getattr(param, "IsReadOnly", False):
                        continue

                    pname = param.Definition.Name
                    # exact match then case-insensitive fallback
                    has_val = pname in param_dict
                    if not has_val:
                        for k in param_dict.keys():
                            if str(k).strip().lower() == pname.lower():
                                pname = k
                                has_val = True
                                break
                    if not has_val:
                        continue

                    new_value = param_dict[pname]
                    sval = "" if new_value is None else str(new_value).strip()

                    st = param.StorageType
                    # Allow clearing text parameters
                    if st == StorageType.String:
                        param.Set(sval)
                        continue

                    # Non-string: skip empty replacement
                    if sval == "":
                        continue

                    if st == StorageType.Integer:
                        param.Set(_to_bool_int(sval))
                        continue
                    if st == StorageType.Double:
                        set_ok = False
                        try:
                            if hasattr(param, "SetValueString"):
                                param.SetValueString(sval)
                                set_ok = True
                        except:
                            set_ok = False
                        if not set_ok:
                            fval = _to_float(sval)
                            if fval is None:
                                continue
                            try:
                                from Autodesk.Revit.DB import UnitUtils
                                if hasattr(param, "GetUnitTypeId"):
                                    uid = param.GetUnitTypeId()
                                    f_internal = UnitUtils.ConvertToInternalUnits(fval, uid)
                                    param.Set(f_internal)
                                    set_ok = True
                            except:
                                set_ok = False
                            if not set_ok:
                                param.Set(fval)
                        continue
                    if st == StorageType.ElementId:
                        # skip automatic ElementId binding
                        continue
                except:
                    pass
        success = True
    except Exception as ex:
        # print("Error updating type parameters: {}".format(str(ex)))
        success = False

    return success

def categorize_types_from_excel(doc, category, excel_data, name_col_index=None, status_col_index=None):
    """
    Categorize types from Excel data into:
    - existing_types: Types that exist in Revit with matching parameters
    - update_types: Types that exist but need parameter updates
    - new_types: Types that don't exist and need creation
    
    Args:
        doc: Revit document
        category: Revit category
        excel_data: List of rows from Excel
        name_col_index: Index of the column containing type names (default: tries to find a column named "Name")
        status_col_index: Index of the column containing status values (default: None, will calculate status)
    
    Returns:
        Dictionary with three lists: 'existing', 'update', 'new'
    """
    if not doc or not category or not excel_data or len(excel_data) < 2:
        return {'existing': [], 'update': [], 'new': []}
    
    result = {
        'existing': [],  # Types that exist with matching parameters
        'update': [],    # Types that exist but parameters differ
        'new': []        # Types that don't exist
    }
    
    try:
        # Get header row (parameter names)
        header_row = excel_data[0]
        
        # Determine the index of the Name column
        if name_col_index is None:
            # Try to find a column named "Name"
            for i, col_name in enumerate(header_row):
                if col_name and str(col_name).strip().lower() == "name":
                    name_col_index = i
                    # print("Found Name column at index {}".format(i))
                    break
            
            # If still not found, default to the first column
            if name_col_index is None:
                name_col_index = 0
                # print("No Name column found, defaulting to column 0")
        
        # print("Using column {} as Name column".format(name_col_index))
        
        # Process each data row
        for row_idx in range(1, len(excel_data)):
            row = excel_data[row_idx]
            if not row or len(row) <= name_col_index or row[name_col_index] is None or str(row[name_col_index]).strip() == "":
                # print("Skipping row {}: Missing or empty name".format(row_idx))
                continue
                
            type_name = str(row[name_col_index]).strip()
            # print("Checking type: {}".format(type_name))
            
            # If status column is provided in Excel, use that value
            if status_col_index is not None and status_col_index < len(row) and row[status_col_index] is not None:
                excel_status = str(row[status_col_index]).strip().lower()
                # print("Using Excel status for {}: {}".format(type_name, excel_status))
                
                # Map Excel status to our categories
                if excel_status in ["existing", "update", "new"]:
                    # Check if type actually exists and parameters match
                    type_exists = check_type_exists_in_category(doc, category, type_name)
                    if type_exists:
                        # Even if marked as "Existing" in Excel, check if parameters need updating
                        diff_params = compare_element_params_with_excel(doc, category, type_name, header_row, row, name_col_index)
                        if diff_params:
                            # Excel says "Existing" but parameters differ, should be "Update"
                            # print("Excel status is {}, but parameters differ - marking as 'update'".format(excel_status))
                            result['update'].append({
                                'name': type_name,
                                'row': row,
                                'diff_params': diff_params,
                                'excel_status': excel_status
                            })
                        else:
                            # Excel says "Existing" and parameters match, correct
                            result['existing'].append({
                                'name': type_name,
                                'row': row,
                                'excel_status': excel_status
                            })
                    else:
                        # Excel says "Existing" but type doesn't exist, should be "New"
                        # print("Excel status is {}, but type doesn't exist - marking as 'new'".format(excel_status))
                        result['new'].append({
                            'name': type_name,
                            'row': row,
                            'excel_status': excel_status
                        })
                        
                elif excel_status in ["update", "needs update", "needsupdate"]:
                    # Check if type actually exists
                    type_exists = check_type_exists_in_category(doc, category, type_name)
                    if type_exists:
                        # Get specific parameter differences
                        diff_params = compare_element_params_with_excel(doc, category, type_name, header_row, row, name_col_index)
                        if diff_params:
                            # Excel says "Update" and parameters differ, correct
                            result['update'].append({
                                'name': type_name,
                                'row': row,
                                'diff_params': diff_params,
                                'excel_status': excel_status
                            })
                        else:
                            # Excel says "Update" but parameters match, should be "Existing"
                            # print("Excel status is {}, but parameters match - marking as 'existing'".format(excel_status))
                            result['existing'].append({
                                'name': type_name,
                                'row': row,
                                'excel_status': excel_status
                            })
                    else:
                        # Excel says "Update" but type doesn't exist, should be "New"
                        # print("Excel status is {}, but type doesn't exist - marking as 'new'".format(excel_status))
                        result['new'].append({
                            'name': type_name,
                            'row': row,
                            'excel_status': excel_status
                        })
                        
                elif excel_status in ["new"]:
                    # Check if type actually exists
                    type_exists = check_type_exists_in_category(doc, category, type_name)
                    if not type_exists:
                        # Excel says "New" and type doesn't exist, correct
                        result['new'].append({
                            'name': type_name,
                            'row': row,
                            'excel_status': excel_status
                        })
                    else:
                        # Excel says "New" but type exists, check if parameters match
                        diff_params = compare_element_params_with_excel(doc, category, type_name, header_row, row, name_col_index)
                        if diff_params:
                            # Type exists but parameters differ, should be "Update"
                            # print("Excel status is {}, but type exists with different parameters - marking as 'update'".format(excel_status))
                            result['update'].append({
                                'name': type_name,
                                'row': row,
                                'diff_params': diff_params,
                                'excel_status': excel_status
                            })
                        else:
                            # Type exists and parameters match, should be "Existing"
                            # print("Excel status is {}, but type exists with matching parameters - marking as 'existing'".format(excel_status))
                            result['existing'].append({
                                'name': type_name,
                                'row': row,
                                'excel_status': excel_status
                            })
                else:
                    # Unknown status, determine based on checks
                    # print("Unknown Excel status: {} - determining status based on checks".format(excel_status))
                    calculate_status = True
            else:
                # No status column in Excel or empty status, determine based on checks
                calculate_status = True
            if 'calculate_status' in locals() and calculate_status:
                # print("Type {} exists: {}".format(type_name, type_exists))
                if type_exists:
                    # Type exists, check parameters
                    diff_params = compare_element_params_with_excel(doc, category, type_name, header_row, row, name_col_index)
                    
                    if diff_params:
                        # Parameters differ, needs update
                        # print("Type needs update: {} ({} parameters differ)".format(
                        #     type_name, len(diff_params)))
                        result['update'].append({
                            'name': type_name,
                            'row': row,
                            'diff_params': diff_params
                        })
                    else:
                        # Parameters match
                        # print("Type matches: {}".format(type_name))
                        result['existing'].append({
                            'name': type_name,
                            'row': row
                        })
                else:
                    # Type doesn't exist, needs creation
                    # print("Type is new: {}".format(type_name))
                    result['new'].append({
                        'name': type_name,
                        'row': row
                    })
    
    except Exception as ex:
        # print("Error categorizing types: {}".format(str(ex)))
        import traceback
        # print(traceback.format_exc())
        pass
    return result
