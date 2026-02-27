# -*- coding: utf-8 -*-
# ===========================
# File: Main_Handler.py - Isolate Element By Parameter Handler
# ===========================
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.UI import IExternalEventHandler
# TaskDialog is not used anymore for debug popups
# from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, TemporaryViewMode, StorageType, ElementId,
    OverrideGraphicSettings, FillPatternElement, FillPatternTarget, Color as DBColor
)
from System.Collections.Generic import List as GenericList

class ExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        self.action = None
        self.view_model = None
        self._is_executing = False
        self._pending_row = None
        self._solid_fill_id = None
        self._last_visible_ids = None
        self._debug_enable = False  # disable all debug popups
        # message returned to caller (script) for simple feedback
        self.message = ""

    def Execute(self, app):
        if self._is_executing:
            return
            
        self._is_executing = True
        try:
            if self.action == "isolate":
                self._isolate_by_parameters(app)
            elif self.action == "unisolate":
                self._clear_isolation(app)
            elif self.action == "select-row":
                self._select_row_elements(app, self._pending_row)
            elif self.action == "color-row":
                self._apply_row_color(app, self._pending_row)
            elif self.action == "select_cad":
                # New: attempt to let user select entities in AutoCAD (lines/text/etc.)
                # The method will set self.message for caller feedback.
                try:
                    self._select_cad_entities(app)
                except Exception as ex:
                    self.message = "AutoCAD selection failed: {}".format(str(ex))
        except Exception as ex:
            print("Error in Execute: {}".format(str(ex)))
        finally:
            self._is_executing = False

    def _debug(self, msg):
        # Debug disabled
        return

    # ===========================
    # Isolate / Unisolate helpers
    # ===========================

    def _isolate_by_parameters(self, app):
        try:
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            view = uidoc.ActiveView
            if not self.view_model:
                print("No view model attached to handler")
                return

            # Collect filters (name, value). Empty value will be ignored
            filters = []
            try:
                for row in self.view_model.ParameterRows:
                    val = getattr(row, 'SelectedValue', None)
                    name = getattr(row, 'ParameterName', None)
                    if name and val not in (None, ""):
                        filters.append((name, val))
            except Exception as ex:
                print("Error reading ParameterRows: {}".format(str(ex)))
                return

            if len(filters) == 0:
                print("No parameter filters selected.")
                return

            # Find matching elements in active view
            elements = list(FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements())
            matched_ids = GenericList[ElementId]()
            try:
                # If AND selected, require all filters to match; otherwise OR
                require_all = getattr(self.view_model, 'IsAndSelected', True)
            except:
                require_all = True

            for el in elements:
                if self._element_matches(el, doc, filters, require_all):
                    matched_ids.Add(el.Id)

            if len(matched_ids) == 0:
                print("No elements match the selected parameter filters.")
                return

            t = Transaction(doc, "Isolate By Parameter")
            t.Start()
            try:
                view.IsolateElementsTemporary(matched_ids)
                t.Commit()
                # cache visible ids and refresh values based on them
                try:
                    self._last_visible_ids = list(matched_ids)
                    if self.view_model:
                        self.view_model.refresh_after_view_change(doc, view, self._last_visible_ids)
                except Exception as ex:
                    print("Refresh after isolate failed: {}".format(str(ex)))
            except Exception as ex:
                print("Failed to isolate elements: {}".format(str(ex)))
                if t and t.HasStarted():
                    t.RollBack()
        except Exception as ex:
            print("Error in isolate: {}".format(str(ex)))

    def _clear_isolation(self, app):
        try:
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            view = uidoc.ActiveView
            t = Transaction(doc, "Clear Temporary Isolation")
            t.Start()
            try:
                view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
                t.Commit()
                # clear cache and refresh with full view
                self._last_visible_ids = None
                if self.view_model:
                    self.view_model.refresh_after_view_change(doc, view, None)
            except Exception as ex:
                print("Failed to clear isolation: {}".format(str(ex)))
                if t and t.HasStarted():
                    t.RollBack()
        except Exception as ex:
            print("Error clearing isolation: {}".format(str(ex)))

    def _select_row_elements(self, app, row):
        """Select all elements in active view that match a single row."""
        try:
            if row is None:
                return
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            view = uidoc.ActiveView

            pname = getattr(row, 'ParameterName', None)
            pval = getattr(row, 'SelectedValue', None)
            if not pname or pval in (None, ""):
                return

            elements = list(FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements())
            ids = GenericList[ElementId]()
            for el in elements:
                if self._element_matches(el, doc, [(pname, pval)], True):
                    ids.Add(el.Id)

            # Select and cache for this row
            uidoc.Selection.SetElementIds(ids)
            try:
                if self.view_model:
                    self.view_model.update_row_cached_ids(row, ids)
            except:
                pass
        except Exception as ex:
            print("Error selecting row elements: {}".format(str(ex)))

    # NEW: apply color override for a single row (like View Specific Element Graphics)
    def _apply_row_color(self, app, row):
        """Apply color overrides to elements for a row; prefer cached ids from Get‑E."""
        try:
            if row is None:
                return
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            view = uidoc.ActiveView

            # Prefer cached ids saved by Get‑E
            cached = getattr(row, 'CachedIds', None)
            ids_to_override = GenericList[ElementId]()
            if cached and len(cached) > 0:
                for i in cached:
                    try:
                        ids_to_override.Add(ElementId(int(i)))
                    except:
                        pass
            else:
                # Fallback: match by parameter/value
                pname = getattr(row, 'ParameterName', None)
                pval = getattr(row, 'SelectedValue', None)
                if not pname or pval in (None, ""):
                    return
                for el in FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements():
                    if self._element_matches(el, doc, [(pname, pval)], True):
                        ids_to_override.Add(el.Id)

            if ids_to_override.Count == 0:
                return

            solid_id = self._get_solid_fill_id(doc)
            if not solid_id or solid_id.IntegerValue == -1:
                return

            # Get color from the row
            wpf_color = getattr(row, 'RowColor', None)
            if wpf_color is None:
                return
            col = self._to_db_color(wpf_color)

            ogs = OverrideGraphicSettings()
            # Apply both new and legacy properties for Revit 2021 compatibility
            try:
                ogs.SetSurfaceForegroundPatternId(solid_id)
                ogs.SetSurfaceForegroundPatternColor(col)
                ogs.SetSurfaceBackgroundPatternId(solid_id)
                ogs.SetSurfaceBackgroundPatternColor(col)
                ogs.SetCutForegroundPatternId(solid_id)
                ogs.SetCutForegroundPatternColor(col)
                ogs.SetCutBackgroundPatternId(solid_id)
                ogs.SetCutBackgroundPatternColor(col)
                for setter in (
                    "SetSurfaceForegroundPatternVisible",
                    "SetSurfaceBackgroundPatternVisible",
                    "SetCutForegroundPatternVisible",
                    "SetCutBackgroundPatternVisible",
                ):
                    try: getattr(ogs, setter)(True)
                    except: pass
            except:
                try: ogs.SetProjectionFillPatternId(solid_id)
                except: pass
                try: ogs.SetProjectionFillColor(col)
                except: pass
                try: ogs.SetCutFillPatternId(solid_id)
                except: pass
                try: ogs.SetCutFillColor(col)
                except: pass
            try: ogs.SetProjectionLineColor(col)
            except: pass

            # Clear previous overrides from this row, then apply new ones
            prev_ids = getattr(row, 'LastOverrideIds', [])
            t = Transaction(doc, "Set color (row)")
            t.Start()
            try:
                if prev_ids:
                    clear_ogs = OverrideGraphicSettings()
                    prev_set = set(int(x) for x in prev_ids)
                    new_set = set([eid.IntegerValue for eid in ids_to_override])
                    for prev in prev_set - new_set:
                        view.SetElementOverrides(ElementId(prev), clear_ogs)

                for eid in ids_to_override:
                    view.SetElementOverrides(eid, ogs)

                t.Commit()
                # remember what we just colored
                try:
                    row.SetLastOverrideIds([eid for eid in ids_to_override])
                except:
                    pass
            except Exception as ex:
                if t.HasStarted(): t.RollBack()
                print("Failed to set overrides: {}".format(str(ex)))
        except Exception as ex:
            print("Error applying row color: {}".format(str(ex)))

    def _element_matches(self, el, doc, filters, require_all=True):
        """Return True if element satisfies the filters.
        require_all=True -> AND logic, False -> OR logic.
        Empty value filters are ignored."""
        try:
            any_match = False
            for (pname, pval) in filters:
                if pval is None or str(pval) == "":
                    # ignore empty value row
                    if require_all:
                        # nothing to check; continue
                        continue
                    else:
                        # for OR nothing to add
                        continue

                # try instance parameter first
                inst_param = el.LookupParameter(pname)
                val = self._get_param_value_as_string(inst_param)
                if val is None:
                    # fallback to type parameter
                    try:
                        typ = doc.GetElement(el.GetTypeId())
                        if typ:
                            type_param = typ.LookupParameter(pname)
                            val = self._get_param_value_as_string(type_param)
                    except:
                        val = None

                wanted = self._normalize(pval)
                have = self._normalize(val)
                is_ok = (have == wanted)

                if require_all:
                    if not is_ok:
                        return False
                else:
                    if is_ok:
                        any_match = True

            return any_match if not require_all else True
        except Exception as ex:
            # On any error, treat as not matching
            return False

    def _get_param_value_as_string(self, param):
        if not param:
            return None
        try:
            # Prefer display string when available
            try:
                s = param.AsValueString()
                if s not in (None, ""):
                    return s
            except:
                pass

            if param.StorageType == StorageType.String:
                return param.AsString()
            elif param.StorageType == StorageType.Integer:
                return str(param.AsInteger())
            elif param.StorageType == StorageType.Double:
                return str(param.AsDouble())
            elif param.StorageType == StorageType.ElementId:
                try:
                    eid = param.AsElementId()
                    return str(eid.IntegerValue)
                except:
                    return ""
            else:
                return ""
        except:
            return None

    def _normalize(self, value):
        try:
            return str(value).strip().lower()
        except:
            return ""

    # Helpers
    def _to_db_color(self, wpf_color):
        try:
            # Revit Color takes byte-range ints 0..255
            return DBColor(int(wpf_color.R), int(wpf_color.G), int(wpf_color.B))
        except Exception as ex:
            print("Color conversion failed: {}".format(str(ex)))
            return DBColor(0, 0, 0)

    def _get_solid_fill_id(self, doc):
        try:
            if self._solid_fill_id and self._solid_fill_id.IntegerValue != -1:
                return self._solid_fill_id

            # Scan FillPatternElements
            for fpe in FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements():
                try:
                    fp = fpe.GetFillPattern()
                    if fp and fp.IsSolidFill:
                        self._solid_fill_id = fpe.Id
                        return self._solid_fill_id
                except:
                    pass

            # Fallback by common localized names
            candidates = ("Solid fill", "Solid Fill", "Solide Füllung", "Plenfyll", "Plný vzor", "Sólido")
            for nm in candidates:
                for tgt in (FillPatternTarget.Drafting, FillPatternTarget.Model):
                    try:
                        fpe = FillPatternElement.GetFillPatternElementByName(doc, tgt, nm)
                        if fpe:
                            self._solid_fill_id = fpe.Id
                            return self._solid_fill_id
                    except:
                        pass

            from Autodesk.Revit.DB import ElementId
            return ElementId.InvalidElementId
        except:
            from Autodesk.Revit.DB import ElementId
            return ElementId.InvalidElementId

    def _select_cad_entities(self, app):
        """Connect to running AutoCAD via COM and prompt user to select entities on screen.
        Collect basic info (ObjectName, Handle) and set self.message. This requires pywin32.
        """
        try:
            try:
                import pythoncom
                import win32com.client
            except Exception as e:
                self.message = "pywin32 not available (win32com). Install pywin32 to enable AutoCAD control."
                return

            # Try get running AutoCAD app first, otherwise try to start one
            acad = None
            try:
                acad = win32com.client.GetActiveObject("AutoCAD.Application")
            except:
                try:
                    acad = win32com.client.Dispatch("AutoCAD.Application")
                except:
                    acad = None

            if acad is None:
                self.message = "AutoCAD application not found."
                return

            # Ensure there's an active document
            try:
                doc = acad.ActiveDocument
            except:
                doc = None
            if doc is None:
                self.message = "No active AutoCAD document."
                return

            # Create a temporary unique selection set name
            import time
            ss_name = "REVIT_TEMP_SS_{}".format(int(time.time()))
            ss = None
            try:
                # If name exists, delete first
                try:
                    existing = doc.SelectionSets.Item(ss_name)
                    try: existing.Delete()
                    except: pass
                except: pass
                ss = doc.SelectionSets.Add(ss_name)
            except Exception as ex:
                self.message = "Failed to create selection set in AutoCAD: {}".format(str(ex))
                return

            try:
                # Prompt user to select on screen. This will block inside AutoCAD until selection done/cancelled.
                try:
                    ss.SelectOnScreen()
                except Exception as ex:
                    # Selection cancelled or error
                    if ss:
                        try: ss.Delete()
                        except: pass
                    self.message = "No entities selected or selection cancelled."
                    return

                count = ss.Count
                if count == 0:
                    try: ss.Delete()
                    except: pass
                    self.message = "No entities selected."
                    return

                items = []
                # Iterate SelectionSet members
                for i in range(count):
                    try:
                        ent = ss.Item(i)
                        objname = getattr(ent, "ObjectName", "<unknown>")
                        handle = getattr(ent, "Handle", None)
                        items.append((objname, str(handle)))
                    except:
                        pass

                # Clean up selection set
                try:
                    ss.Delete()
                except:
                    pass

                # Build concise message (limit length)
                preview = ", ".join(["{}({})".format(n.split('.')[-1], h) for n, h in items[:20]])
                more = " ..." if len(items) > 20 else ""
                self.message = "Selected {} entities in AutoCAD: {}{}".format(len(items), preview, more)
            except Exception as ex:
                try:
                    if ss: ss.Delete()
                except: pass
                self.message = "Error during AutoCAD selection: {}".format(str(ex))
                return
        except Exception as ex:
            self.message = "Unexpected error: {}".format(str(ex))
            return

    def GetName(self):
        return "Isolate Element By Parameter Handler"