# -*- coding: utf-8 -*-
# ===========================
# File: Main_Handler.py - Duplicate Element Type Handler
# ===========================

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import IExternalEventHandler
from Autodesk.Revit.DB import Transaction, ElementId, StorageType  # <-- added StorageType

class ExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        self.source_type = None
        self.new_type_name = None
        self.excel_data = None
        self.excel_data_rows = None  # <-- support the ViewModel's ExcelDataRows
        self.action = None
        self.view_model = None
        self._is_executing = False

    def Execute(self, app):
        if self._is_executing:
            return
            
        self._is_executing = True
        try:
            if self.action == "create_duplicate_type":
                self._create_duplicate_type(app)
        except Exception as ex:
            # print("Error in Execute: {}".format(str(ex)))
            pass
        finally:
            self._is_executing = False

    def _create_duplicate_type(self, app):
        """Create duplicate element type"""
        t = None
        try:
            doc = app.ActiveUIDocument.Document
            
            if not doc or not self.source_type or not self.new_type_name:
                # print("Missing required data for duplication")
                return

            # Start transaction
            t = Transaction(doc, "Duplicate Element Type")
            t.Start()

            try:
                # Duplicate the type. Duplicate may return an ElementId (or element depending on API),
                # so handle multiple possible return shapes.
                dup_result = None
                try:
                    dup_result = self.source_type.Duplicate(self.new_type_name)
                except Exception as ex:
                    # print("Duplicate call failed: {}".format(str(ex)))
                    dup_result = None

                # Normalize dup_result -> try to obtain the Element instance (new_type_elem)
                new_type_elem = None
                try:
                    if dup_result is None:
                        new_type_elem = None
                    else:
                        # If it's already ElementId
                        if isinstance(dup_result, ElementId):
                            new_type_elem = doc.GetElement(dup_result)
                        else:
                            # If object has IntegerValue (ElementId-like), use it
                            iv = getattr(dup_result, 'IntegerValue', None)
                            if iv is not None:
                                try:
                                    new_type_elem = doc.GetElement(ElementId(int(iv)))
                                except:
                                    new_type_elem = None
                            else:
                                # Try doc.GetElement directly (some runtimes accept ElementId-like)
                                try:
                                    new_type_elem = doc.GetElement(dup_result)
                                except:
                                    # Fallback: maybe Duplicate returned an element object already
                                    new_type_elem = dup_result
                except Exception:
                    new_type_elem = dup_result

                if new_type_elem:
                    # print("Successfully created new type: {}".format(self.new_type_name))
                    
                    # Apply Excel data if available (supports ViewModel ExcelDataRows)
                    try:
                        excel_rows = None
                        if self.excel_data_rows:
                            excel_rows = self.excel_data_rows
                        elif self.excel_data:
                            # backward compatibility
                            excel_rows = self.excel_data

                        if excel_rows and hasattr(excel_rows, 'Count') and excel_rows.Count > 0:
                            self._apply_excel_data_to_type(new_type_elem, excel_rows)
                            
                    except Exception as ex:
                        # print("Error applying excel data: {}".format(str(ex)))
                        pass
                    
                    t.Commit()
                    # print("Type duplication completed successfully")
                else:
                    # print("Failed to duplicate type")
                    t.RollBack()
            except Exception as ex:
                # print("Error in duplication transaction: {}".format(str(ex)))
                if t and t.HasStarted():
                    t.RollBack()
        except Exception as ex:
            # print("Error duplicating type: {}".format(str(ex)))
            if t and t.HasStarted():
                t.RollBack()

    def _apply_excel_data_to_type(self, element_type, excel_rows):
        """Apply Excel data to the new element type.
        The first row is treated as parameter names (headers)
        Subsequent rows contain parameter values for duplicate element types
        """
        try:
            if not excel_rows or excel_rows.Count == 0:
                return

            # Always treat first row as parameter names (headers)
            if excel_rows.Count < 2:
                # print("Excel data must have at least two rows: header row and value row")
                return

            # Get the header row (parameter names)
            header_row_item = excel_rows[0]
            header_row = getattr(header_row_item, 'RowData', None)
            if header_row is None or not any(header_row):
                # print("Header row is empty or invalid")
                return
            
            # Convert header row to strings
            header_names = [str(c).strip().lower() if c is not None else "" for c in header_row]
            
            # Build parameter lookup by name (lowercase)
            type_parameters = {}
            for param in element_type.Parameters:
                if param.Definition and param.Definition.Name:
                    type_parameters[param.Definition.Name.strip().lower()] = param
            
            # For now, we'll use the second row as parameter values
            # In a future enhancement, you could apply multiple rows to create multiple duplicates
            if excel_rows.Count > 1:
                data_row_item = excel_rows[1]
                data_row = getattr(data_row_item, 'RowData', [])
                
                # Map parameters by header names
                for idx, param_name in enumerate(header_names):
                    if not param_name:
                        continue
                    
                    if param_name in type_parameters and idx < len(data_row):
                        param = type_parameters[param_name]
                        if not param.IsReadOnly:
                            try:
                                val = data_row[idx]
                                val_str = "" if val is None else str(val)
                                self._set_parameter_value(param, val_str)
                                # print("Set parameter '{}' = '{}'".format(param.Definition.Name, val_str))
                            except Exception as ex:
                                # print("Error setting parameter '{}': {}".format(param.Definition.Name, str(ex)))
                                pass
        except Exception as ex:
            # print("Error applying Excel data: {}".format(str(ex)))
            import traceback
            traceback.print_exc()

    def _set_parameter_value(self, param, value_string):
        """Set parameter value from string"""
        try:
            # Normalize string
            val = "" if value_string is None else value_string.strip()

            # Use StorageType imported from Autodesk.Revit.DB
            if param.StorageType == StorageType.String:
                param.Set(val)
            elif param.StorageType == StorageType.Integer:
                try:
                    int_value = int(float(val)) if val not in ("", None) else 0
                    param.Set(int_value)
                except:
                    # print("Cannot convert '{}' to integer for parameter '{}'".format(
                    #     value_string, param.Definition.Name))
                    pass
            elif param.StorageType == StorageType.Double:
                try:
                    double_value = float(val) if val not in ("", None) else 0.0
                    param.Set(double_value)
                except:
                    # print("Cannot convert '{}' to double for parameter '{}'".format(
                    #     value_string, param.Definition.Name))
                    pass
            elif param.StorageType == StorageType.ElementId:
                # Thử chuyển giá trị thành integer -> ElementId. Nếu không được, bỏ qua.
                try:
                    if val == "":
                        # dùng ElementId(-1) làm invalid fallback
                        param.Set(ElementId(-1))
                    else:
                        int_val = int(float(val))
                        param.Set(ElementId(int_val))
                except Exception:
                    # Không thể convert, có thể cần tìm ElementId theo tên (ngoài phạm vi hiện tại)
                    # print("Cannot convert '{}' to ElementId for parameter '{}'".format(
                    #     value_string, param.Definition.Name))
                    pass
            else:
                # For other types, try to set as string where supported
                try:
                    param.Set(val)
                except Exception:
                    # Some storage types need special handling which is out of scope here
                    # print("Unable to set parameter '{}' with provided value '{}' (unsupported storage type)".format(
                    #     param.Definition.Name, value_string))
                    pass
        except Exception as ex:
            # print("Error setting parameter value: {}".format(str(ex)))
            pass

    def GetName(self):
        return "Duplicate Element Type Handler"