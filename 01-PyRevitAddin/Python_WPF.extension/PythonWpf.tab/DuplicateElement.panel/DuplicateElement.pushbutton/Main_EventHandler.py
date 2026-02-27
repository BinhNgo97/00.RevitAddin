# -*- coding: utf-8 -*-
# ===========================
# File: Main_EventHandler.py - Event handler for Duplicate Element Type
# ===========================

import clr
import sys
import os

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    Transaction, ElementType, FamilySymbol,
    ElementId, BuiltInParameter, StorageType
)
from Autodesk.Revit.UI import (
    IExternalEventHandler, TaskDialog
)

class ExternalEventHandler(IExternalEventHandler):
    """
    Handler for external events to perform operations that need to be 
    executed within a valid Revit API context
    """
    def __init__(self):
        self.source_type = None  # ElementType to duplicate
        self.new_type_name = None  # Name for the new type
        self.new_types_info = None  # List of type infos for batch creation
        self.parameter_dict = None  # Dict of parameter values to set
        self.excel_data_rows = None  # Excel data rows
        self.header_row = None  # Header row from Excel
        self.action = None  # Action to perform
        self.element_type = None  # ElementType to update
        self.param_dict = None  # Dict of parameters to update

    def Execute(self, app):
        """
        Execute method required by IExternalEventHandler interface
        Called when the external event is raised
        """
        try:
            # Get active document
            doc = app.ActiveUIDocument.Document
            if not doc:
                # print("No active document")
                return
            if self.action == "create_duplicate_type":
                self._create_duplicate_type(doc)
            elif self.action == "create_multiple_types":
                self._create_multiple_types(doc)
            elif self.action == "update_element_parameters":
                self._update_element_parameters(doc)
            else:
                print("Unknown action: {}".format(self.action))
                
        except Exception as ex:
            # print("Error in Execute: {}".format(str(ex)))
            import traceback
            # print(traceback.format_exc())
            pass
        finally:
            # Clear state
            self.action = None
    
    def GetName(self):
        """Return the name of the external event handler"""
        return "DuplicateElementTypeHandler"
    
    def _create_duplicate_type(self, doc):
        """Create a duplicate element type with optional parameter values"""
        if not self.source_type or not self.new_type_name:
            print("Missing required parameters for duplicate type creation")
            return
        try:
            with Transaction(doc, "Duplicate Element Type") as trans:
                trans.Start()
                
                # Create duplicate type (may return ElementId or Element)
                dup_result = None
                try:
                    dup_result = self.source_type.Duplicate(self.new_type_name)
                except Exception as ex:
                    # print("Duplicate call failed: {}".format(str(ex)))
                    dup_result = None

                # Normalize to element instance
                new_type = None
                try:
                    if dup_result is None:
                        new_type = None
                    else:
                        from Autodesk.Revit.DB import ElementId
                        # If ElementId returned
                        if isinstance(dup_result, ElementId):
                            new_type = doc.GetElement(dup_result)
                        else:
                            # If element-like returned, use directly
                            # Avoid passing element into GetElement (which expects ElementId or Reference)
                            new_type = dup_result
                except Exception as ex_norm:
                    # print("Normalization error: {}".format(str(ex_norm)))
                    new_type = None
                
                # Update parameters if provided
                result = None
                if new_type and self.parameter_dict:
                    result = self._set_parameters(new_type, self.parameter_dict)
                
                trans.Commit()

            # Notify result
            msg = "Created new type: {}\n".format(self.new_type_name)
            if result:
                msg += "- Parameters set: {set}\n- Skipped: {skipped}\n- Failed: {failed}".format(
                    set=result.get('set', 0), skipped=result.get('skipped', 0), failed=result.get('failed', 0)
                )
                # Add skip details if available
                details = []
                if result.get('skipped_readonly'):
                    details.append("Readonly: {}".format(", ".join(result['skipped_readonly'][:6])))
                if result.get('skipped_notfound'):
                    details.append("Not found: {}".format(", ".join(result['skipped_notfound'][:6])))
                if result.get('skipped_empty'):
                    details.append("Empty: {}".format(", ".join(result['skipped_empty'][:6])))
                if result.get('skipped_elementid'):
                    details.append("ElementId: {}".format(", ".join(result['skipped_elementid'][:6])))
                if result.get('failed_double_parse'):
                    details.append("Number parse failed: {}".format(", ".join(result['failed_double_parse'][:6])))
                if details:
                    msg += "\n\nSkipped details:\n" + "\n".join(details)
                if result.get('errors'):
                    msg += "\n\nErrors:\n" + "\n".join(result['errors'][:10])
            TaskDialog.Show("Duplicate Type", msg)  # keep success popup
        except Exception as ex:
            # print("Error creating duplicate type: {}".format(str(ex)))
            try:
                # TaskDialog.Show("Duplicate Type", "Error: {}".format(str(ex)))  # disabled popup
                pass
            except:
                pass

    def _create_multiple_types(self, doc):
        """Create multiple types from a list of type infos"""
        if not self.source_type or not self.new_types_info or not self.header_row:
            print("Missing required parameters for batch type creation")
            return
        try:
            total = len(self.new_types_info)
            created = 0
            set_total = 0
            skipped_total = 0
            failed_total = 0
            errors = []

            for type_info in self.new_types_info:
                type_name = type_info['name']
                excel_row = type_info['row']
                with Transaction(doc, "Create Type - {}".format(type_name)) as trans:
                    trans.Start()
                    # Duplicate the source type
                    dup_result = None
                    try:
                        dup_result = self.source_type.Duplicate(type_name)
                    except Exception as ex_dup:
                        dup_result = None
                        # errors.append("{}: Duplicate failed: {}".format(type_name, str(ex_dup)))
                        pass

                    # Normalize dup_result to an element instance
                    new_type = None
                    try:
                        if dup_result is None:
                            new_type = None
                        else:
                            # If Duplicate returned ElementId
                            if isinstance(dup_result, ElementId):
                                new_type = doc.GetElement(dup_result)
                            else:
                                # If Duplicate returned an element already (e.g., FloorType), use it directly
                                new_type = dup_result
                    except Exception as ex_norm:
                        # errors.append("{}: Normalize failed: {}".format(type_name, str(ex_norm)))
                        new_type = None

                    if new_type:
                        # Build parameter dict from Excel row
                        param_dict = {}
                        for i in range(1, min(len(self.header_row), len(excel_row))):
                            if self.header_row[i] and str(self.header_row[i]).strip():
                                param_name = str(self.header_row[i]).strip()
                                param_dict[param_name] = excel_row[i]
                        # Set parameters
                        result = self._set_parameters(new_type, param_dict)
                        set_total += result.get('set', 0)
                        skipped_total += result.get('skipped', 0)
                        failed_total += result.get('failed', 0)
                        if result.get('errors'):
                            errors.extend(["{}: {}".format(type_name, e) for e in result['errors']])
                        created += 1
                        # print("Created type: {}".format(type_name))
                        pass
                    else:
                        # errors.append("{}: Duplicate returned no element".format(type_name))
                        pass

                    trans.Commit()

            # Notify result summary
            msg = "Created {}/{} types.\nParameters set: {} | Skipped: {} | Failed: {}".format(
                created, total, set_total, skipped_total, failed_total
            )
            if errors:
                msg += "\n\nTop errors:\n" + "\n".join(errors[:10])
            # TaskDialog.Show("Create Multiple Types", msg)  # disabled popup
            print(msg)
        except Exception as ex:
            # print("Error creating multiple types: {}".format(str(ex)))
            try:
                # TaskDialog.Show("Create Multiple Types", "Error: {}".format(str(ex)))  # disabled popup
                pass
            except:
                pass

    def _update_element_parameters(self, doc):
        """Update parameters of an existing element type"""
        if not self.element_type or not self.param_dict:
            print("Missing required parameters for update")
            return
        try:
            with Transaction(doc, "Update Element Type Parameters") as trans:
                trans.Start()
                result = self._set_parameters(self.element_type, self.param_dict)
                trans.Commit()
            # Notify result
            tname = getattr(self.element_type, "Name", "Unknown")
            msg = "Updated type: {}\n- Parameters set: {}\n- Skipped: {}\n- Failed: {}".format(
                tname, result.get('set', 0), result.get('skipped', 0), result.get('failed', 0)
            )
            details = []
            if result.get('skipped_readonly'):
                details.append("Readonly: {}".format(", ".join(result['skipped_readonly'][:6])))
            if result.get('skipped_notfound'):
                details.append("Not found: {}".format(", ".join(result['skipped_notfound'][:6])))
            if result.get('skipped_empty'):
                details.append("Empty: {}".format(", ".join(result['skipped_empty'][:6])))
            if result.get('skipped_elementid'):
                details.append("ElementId: {}".format(", ".join(result['skipped_elementid'][:6])))
            if result.get('failed_double_parse'):
                details.append("Number parse failed: {}".format(", ".join(result['failed_double_parse'][:6])))
            if details:
                msg += "\n\nSkipped details:\n" + "\n".join(details)
            if result.get('errors'):
                msg += "\n\nErrors:\n" + "\n".join(result['errors'][:10])
            TaskDialog.Show("Update Type Parameters", msg)  # keep success popup
        except Exception as ex:
            # print("Error updating element parameters: {}".format(str(ex)))
            try:
                # TaskDialog.Show("Update Type Parameters", "Error: {}".format(str(ex)))  # disabled popup
                pass
            except:
                pass

    def _set_parameters(self, element, param_dict):
        """Set parameters on an element from a dictionary. Returns result summary."""
        result = {
            'set': 0, 'skipped': 0, 'failed': 0, 'errors': [],
            'skipped_readonly': [], 'skipped_notfound': [], 'skipped_empty': [],
            'failed_double_parse': [], 'skipped_elementid': []
        }
        if not element or not param_dict:
            return result

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

        for param in element.Parameters:
            try:
                if not param or not param.Definition or not param.Definition.Name:
                    continue

                if getattr(param, "IsReadOnly", False):
                    result['skipped'] += 1
                    try:
                        result['skipped_readonly'].append(param.Definition.Name)
                    except:
                        pass
                    continue

                pname = param.Definition.Name
                # exact name first, then case-insensitive
                has_val = pname in param_dict
                if not has_val:
                    for k in param_dict.keys():
                        if str(k).strip().lower() == pname.lower():
                            pname = k
                            has_val = True
                            break
                if not has_val:
                    result['skipped'] += 1
                    try:
                        result['skipped_notfound'].append(param.Definition.Name)
                    except:
                        pass
                    continue

                new_value = param_dict[pname]
                sval = "" if new_value is None else str(new_value).strip()

                from Autodesk.Revit.DB import StorageType as ST
                st = param.StorageType

                # Allow clearing text parameters to empty string
                if st == ST.String:
                    try:
                        param.Set(sval)
                        result['set'] += 1
                    except Exception as ex_set:
                        result['failed'] += 1
                        result['errors'].append("Param '{}': {}".format(
                            getattr(param.Definition, "Name", "?"), str(ex_set)))
                    continue

                # For non-string params: skip truly empty values
                if sval == "":
                    result['skipped'] += 1
                    try:
                        result['skipped_empty'].append(param.Definition.Name)
                    except:
                        pass
                    continue

                if st == ST.Integer:
                    ival = _to_bool_int(sval)
                    param.Set(ival)
                    result['set'] += 1
                    continue

                if st == ST.Double:
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
                            result['skipped'] += 1
                            try:
                                result['failed_double_parse'].append(param.Definition.Name)
                            except:
                                pass
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
                    result['set'] += 1
                    continue

                if st == ST.ElementId:
                    result['skipped'] += 1
                    try:
                        result['skipped_elementid'].append(param.Definition.Name)
                    except:
                        pass
                    continue

                # Fallback
                try:
                    if hasattr(param, "SetValueString"):
                        param.SetValueString(sval)
                        result['set'] += 1
                    else:
                        result['skipped'] += 1
                except Exception as ex_set:
                    result['failed'] += 1
                    result['errors'].append("Param '{}': {}".format(
                        getattr(param.Definition, "Name", "?"), str(ex_set)))
            except Exception as ex:
                result['failed'] += 1
                result['errors'].append("Param '{}': {}".format(
                    getattr(param.Definition, "Name", "?"), str(ex)))
        return result
