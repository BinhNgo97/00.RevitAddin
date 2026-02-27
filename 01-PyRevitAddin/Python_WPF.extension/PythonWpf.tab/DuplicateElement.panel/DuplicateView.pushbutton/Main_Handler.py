# -*- coding: utf-8 -*-
# ===========================
# File: Main_Handler.py - Isolate Element By Parameter Handler
# ===========================

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import IExternalEventHandler
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, TemporaryViewMode, StorageType, ElementId, View, ViewDuplicateOption
)
from Autodesk.Revit.DB import SubTransaction
from System.Collections.Generic import List as GenericList

class ExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        self.action = None
        self.view_model = None
        self._is_executing = False
        self.duplicate_mode = "basic"  # basic | detail | dependent
        # for apply_template
        self.target_view_id = None
        self.template_id = None
        self.template_name = None
        self.target_view_ids = None  # list for multi-apply

    def Execute(self, app):
        if self._is_executing:
            return
            
        self._is_executing = True
        try:
            if self.action == "isolate":
                self._isolate_by_parameters(app)
            elif self.action == "unisolate":
                self._clear_isolation(app)
            elif self.action == "duplicate":
                self._duplicate_selected_views(app)
            elif self.action == "rename":
                self._rename_selected_views(app)
            elif self.action == "apply_template":
                self._apply_view_template(app)
            elif self.action == "apply_template_multi":
                self._apply_view_template_multi(app)
            # New: load selected views into grid
            elif self.action == "load_selection":
                self._load_selected_views(app)
        except Exception as ex:
            print("Error in Execute: {}".format(str(ex)))
        finally:
            self._is_executing = False

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
                    pname = getattr(row, 'ParameterName', None)
                    pval = getattr(row, 'SelectedValue', None)
                    if pname and (pval is not None):
                        filters.append((pname, pval))
            except Exception as ex:
                print("Error reading ParameterRows: {}".format(str(ex)))
                return

            if len(filters) == 0:
                print("No parameter filters selected.")
                return

            # Find matching elements in active view
            elements = list(FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements())
            matched_ids = GenericList[ElementId]()
            require_all = True
            try:
                # If OR selected, match any
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
            except Exception as ex:
                print("Failed to isolate elements: {}".format(str(ex)))
                if t and t.HasStarted():
                    t.RollBack()

        except Exception as ex:
            print("Error in isolate: {}".format(str(ex)))

    def _clear_isolation(self, app):
        try:
            uidoc = app.ActiveUIDocument
            view = uidoc.ActiveView
            t = Transaction(uidoc.Document, "Clear Temporary Isolation")
            t.Start()
            try:
                view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
                t.Commit()
            except Exception as ex:
                print("Failed to clear isolation: {}".format(str(ex)))
                if t and t.HasStarted():
                    t.RollBack()
        except Exception as ex:
            print("Error clearing isolation: {}".format(str(ex)))

    def _element_matches(self, el, doc, filters, require_all=True):
        """Return True if element satisfies the filters.
        require_all=True -> AND logic, False -> OR logic.
        Empty value filters are ignored."""
        try:
            any_match = False
            for (pname, pval) in filters:
                if pval is None or str(pval) == "":
                    # ignore empty value filter
                    continue

                inst_param = el.LookupParameter(pname)
                val = self._get_param_value_as_string(inst_param)
                if val is None:
                    # try type parameter
                    try:
                        type_id = el.GetTypeId()
                        if type_id and type_id.IntegerValue != -1:
                            eltype = doc.GetElement(type_id)
                            type_param = eltype.LookupParameter(pname) if eltype else None
                            val = self._get_param_value_as_string(type_param)
                    except:
                        val = None

                # Compare as string (case-insensitive, trimmed)
                if val is None:
                    if require_all:
                        return False
                    else:
                        continue
                match = (self._normalize(val) == self._normalize(pval))

                if require_all and not match:
                    return False
                if not require_all and match:
                    return True
                any_match = any_match or match

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
                disp = param.AsValueString()
                if disp not in (None, ""):
                    return disp
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
                    if eid and eid.IntegerValue != -1:
                        doc = __revit__.ActiveUIDocument.Document
                        el = doc.GetElement(eid)
                        if el is not None and hasattr(el, 'Name'):
                            return el.Name or str(eid.IntegerValue)
                        return str(eid.IntegerValue)
                    return ""
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

    def GetName(self):
        return "Isolate Element By Parameter Handler"

    # ---------------- Duplicate ----------------
    def _duplicate_selected_views(self, app):
        try:
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            vm = self.view_model
            if vm is None:
                return

            # Map mode to API option
            opt = ViewDuplicateOption.Duplicate
            mode = (self.duplicate_mode or "basic").lower()
            if mode == "detail":
                opt = ViewDuplicateOption.WithDetailing
            elif mode == "dependent":
                opt = ViewDuplicateOption.AsDependent

            # Collect selected views (exclude templates)
            sel_ids = list(uidoc.Selection.GetElementIds())
            src_views = []
            for eid in sel_ids:
                el = doc.GetElement(eid)
                if isinstance(el, View) and not el.IsTemplate:
                    src_views.append(el)
            if not src_views:
                print("Select one or more views in Revit.")
                return

            t = Transaction(doc, "Duplicate Views")
            t.Start()
            new_views = []
            try:
                for v in src_views:
                    try:
                        nid = v.Duplicate(opt)
                        nv = doc.GetElement(nid)
                        if nv:
                            new_views.append(nv)
                    except Exception as ex:
                        # skip views that cannot be duplicated with chosen option
                        pass
                t.Commit()
            except Exception as ex:
                if t and t.HasStarted():
                    t.RollBack()
                print("Duplicate failed: {}".format(str(ex)))
                return

            # Force regen to refresh view/template relationships after duplication
            try:
                doc.Regenerate()
            except:
                pass

            # Push to VM
            for nv in new_views:
                tmpl = "none"
                try:
                    vtid = nv.ViewTemplateId
                    if vtid and isinstance(vtid, ElementId) and vtid.IntegerValue > 0:
                        vt = doc.GetElement(vtid)
                        tmpl = vt.Name if vt else "none"
                except:
                    tmpl = "none"
                try:
                    vm.add_duplicated_view(nv.Id, nv.Name, tmpl)
                    # also set template choices for this view type
                    pairs = self._collect_templates_for_view(doc, nv)
                    try:
                        vm.set_view_templates(nv.Id, pairs)
                    except:
                        pass
                except Exception as ex:
                    pass

        except Exception as ex:
            print("Error in _duplicate_selected_views: {}".format(str(ex)))

    def _are_viewtypes_compatible(self, template_vtype, target_vtype):
        # Groups follow Revit UI filter groupings: Plans, Elevations/Sections/Details, 3D/Walkthrough, etc.
        try:
            group_plans = set([ViewType.FloorPlan, ViewType.EngineeringPlan, ViewType.AreaPlan, ViewType.CeilingPlan])
            group_esd = set([ViewType.Elevation, ViewType.Section, ViewType.Detail])  # Elevations, Sections, Detail Views
            group_3d = set([ViewType.ThreeD, ViewType.Walkthrough])  # 3D Views, Walkthroughs
            # Single-type groups
            singles = [
                ViewType.DraftingView,
                ViewType.Legend,
                ViewType.DrawingSheet,
                ViewType.Rendering,
                ViewType.Schedule
            ]

            # Same type always compatible
            if template_vtype == target_vtype:
                return True

            # Plans group cross-compatibility
            if template_vtype in group_plans and target_vtype in group_plans:
                return True

            # Elevation/Section/Detail cross-compatibility
            if template_vtype in group_esd and target_vtype in group_esd:
                return True

            # 3D/Walkthrough cross-compatibility
            if template_vtype in group_3d and target_vtype in group_3d:
                return True

            # Singles: only same type allowed
            if template_vtype in singles or target_vtype in singles:
                return False

            # Fallback: no
            return False
        except:
            return False

    def _can_apply_template(self, doc, view, template_id):
        """Safely test if a template can be applied to a view by using a SubTransaction and rolling back."""
        try:
            if view is None or template_id is None:
                return False
            if hasattr(view, 'IsTemplate') and view.IsTemplate:
                return False
            st = SubTransaction(doc)
            st.Start()
            try:
                original = view.ViewTemplateId
                view.ViewTemplateId = template_id
                # if no exception, consider applicable
                st.RollBack()
                return True
            except:
                st.RollBack()
                return False
        except:
            return False

    def _collect_templates_for_view(self, doc, view):
        """Return list of (ElementId, Name) for templates applicable to the given view, including 'None' first."""
        try:
            # Ensure latest document state and a fresh view handle
            try:
                doc.Regenerate()
                view = doc.GetElement(view.Id)
            except:
                pass

            results = []
            try:
                invalid_id = ElementId.InvalidElementId
            except:
                invalid_id = ElementId(-1)
            results.append((invalid_id, "None"))

            # Collect all template views in document
            try:
                collector = FilteredElementCollector(doc).OfClass(View)
                templates = [v for v in collector if hasattr(v, 'IsTemplate') and v.IsTemplate]
            except:
                templates = []

            applicable = []

            # 1) Preferred: use API to get valid template ids if available
            try:
                get_valid = getattr(view, 'GetValidViewTemplateIds', None)
                if get_valid is not None:
                    valid_ids = list(get_valid())  # ICollection<ElementId>
                    valid_set = set([vid.IntegerValue for vid in valid_ids if hasattr(vid, 'IntegerValue')])
                    for tview in templates:
                        try:
                            if tview.Id.IntegerValue in valid_set:
                                applicable.append((tview.Id, tview.Name))
                        except:
                            pass
            except:
                pass

            # 2) Fallback: probe applicability inside a temporary transaction using SubTransactions
            if not applicable:
                probe_tx = None
                try:
                    probe_tx = Transaction(doc, "Probe View Template Applicability")
                    probe_tx.Start()
                    for tview in templates:
                        st = SubTransaction(doc)
                        st.Start()
                        ok = False
                        try:
                            fresh = doc.GetElement(view.Id)
                            # attempt assignment; if no exception, it's applicable
                            fresh.ViewTemplateId = tview.Id
                            ok = True
                        except:
                            ok = False
                        finally:
                            # always roll back subtransaction
                            try:
                                st.RollBack()
                            except:
                                pass
                        if ok:
                            applicable.append((tview.Id, tview.Name))
                except:
                    pass
                finally:
                    # roll back the outer probing transaction to leave model unchanged
                    try:
                        if probe_tx and probe_tx.HasStarted():
                            probe_tx.RollBack()
                    except:
                        pass

            # 3) Secondary: broaden by cross-viewtype compatibility (optional)
            try:
                vtype = getattr(view, 'ViewType', None)
                for tview in templates:
                    try:
                        tvt = getattr(tview, 'ViewType', None)
                        pair = (tview.Id, tview.Name)
                        if self._are_viewtypes_compatible(tvt, vtype) and pair not in applicable:
                            applicable.append(pair)
                    except:
                        pass
            except:
                pass

            # Sort by name, keep "None" first
            applicable = sorted(applicable, key=lambda p: (p[1] or "").lower())
            return results + applicable
        except Exception as ex:
            # On any error, at least return 'None'
            try:
                invalid_id = ElementId.InvalidElementId
            except:
                invalid_id = ElementId(-1)
            return [(invalid_id, "None")]

    # ---------------- Apply Template (single) ----------------
    def _apply_view_template(self, app):
        try:
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            vm = self.view_model
            if vm is None or self.target_view_id is None:
                return

            vid = self.target_view_id
            tid = self.template_id
            tname = self.template_name or "none"

            v = doc.GetElement(vid)
            if v is None or not isinstance(v, View):
                return

            # safe invalid id
            try:
                invalid_id = ElementId.InvalidElementId
            except:
                invalid_id = ElementId(-1)

            t = Transaction(doc, "Set View Template")
            t.Start()
            try:
                if tid is None:
                    v.ViewTemplateId = invalid_id
                else:
                    # IronPython sometimes wraps ElementId; compare by IntegerValue
                    if (hasattr(tid, 'IntegerValue') and tid.IntegerValue == -1) or tid == invalid_id:
                        v.ViewTemplateId = invalid_id
                    else:
                        v.ViewTemplateId = tid
                t.Commit()
            except Exception as ex:
                if t and t.HasStarted():
                    t.RollBack()
                print("Failed to apply template: {}".format(str(ex)))
                return

            # update VM
            try:
                vm.update_view_template(v.Id, tname if tname != "None" else "none")
            except:
                pass
        except Exception as ex:
            print("Error in _apply_view_template: {}".format(str(ex)))

    # ---------------- Apply Template (multi) ----------------
    def _apply_view_template_multi(self, app):
        try:
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            vm = self.view_model
            ids = list(self.target_view_ids or [])
            if vm is None or not ids:
                return

            # safe invalid id
            try:
                invalid_id = ElementId.InvalidElementId
            except:
                invalid_id = ElementId(-1)

            tid = self.template_id
            tname = self.template_name or "none"

            t = Transaction(doc, "Set View Template (Multiple)")
            t.Start()
            applied_ids = []
            try:
                for vid in ids:
                    try:
                        v = doc.GetElement(vid)
                        if v is None or not isinstance(v, View) or getattr(v, 'IsTemplate', False):
                            continue
                        # attempt assignment; skip on failure
                        if tid is None or (hasattr(tid, 'IntegerValue') and tid.IntegerValue == -1) or tid == invalid_id:
                            v.ViewTemplateId = invalid_id
                        else:
                            v.ViewTemplateId = tid
                        applied_ids.append(vid)
                    except:
                        # ignore incompatible views
                        pass
                t.Commit()
            except Exception as ex:
                if t and t.HasStarted():
                    t.RollBack()
                print("Failed to apply template (multi): {}".format(str(ex)))
                return

            # Regenerate and update VM rows
            try:
                doc.Regenerate()
            except:
                pass
            try:
                for vid in applied_ids:
                    vm.update_view_template(vid, tname if tname != "None" else "none")
            except:
                pass
        except Exception as ex:
            print("Error in _apply_view_template_multi: {}".format(str(ex)))

    def _rename_selected_views(self, app):
        try:
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            vm = self.view_model
            if vm is None:
                return

            find_txt = (getattr(vm, 'FindText', "") or "").strip()
            repl_txt = getattr(vm, 'ReplaceText', "") or ""
            prefix_txt = getattr(vm, 'PrefixText', "") or ""
            suffix_txt = getattr(vm, 'SuffixText', "") or ""
            items = list(getattr(vm, 'SelectedViewItems', []) or [])

            if not items or not (find_txt or prefix_txt or suffix_txt):
                print("Nothing to rename.")
                return

            t = Transaction(doc, "Rename Views")
            t.Start()
            try:
                renamed, failed = 0, 0
                for item in items:
                    try:
                        vid = getattr(item, 'ElementId', None)
                        v = doc.GetElement(vid)
                        if v is None:
                            continue
                        cur = v.Name or ""
                        new_name = cur

                        if find_txt:
                            # Replace mode (repl_txt may be empty to delete)
                            new_name = new_name.replace(find_txt, repl_txt)
                        else:
                            # Prefix/Suffix mode
                            if prefix_txt:
                                new_name = prefix_txt + new_name
                            if suffix_txt:
                                new_name = new_name + suffix_txt

                        if new_name and new_name != cur:
                            v.Name = new_name
                            try:
                                item.ViewName = new_name
                            except:
                                pass
                            renamed += 1
                    except:
                        failed += 1
                        # continue
                t.Commit()
                if renamed == 0 and failed == 0:
                    print("No view names were changed.")
                elif failed > 0:
                    print("Renamed: {}. Failed: {} (duplicate names or invalid characters).".format(renamed, failed))
            except Exception as ex:
                if t and t.HasStarted():
                    t.RollBack()
                print("Rename failed: {}".format(str(ex)))
        except Exception as ex:
            print("Error in _rename_selected_views: {}".format(str(ex)))

    # New: load selected views (no duplication)
    def _load_selected_views(self, app):
        try:
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            vm = self.view_model
            if vm is None:
                return

            sel_ids = list(uidoc.Selection.GetElementIds())
            views = []
            for eid in sel_ids:
                el = doc.GetElement(eid)
                if isinstance(el, View) and not getattr(el, 'IsTemplate', False):
                    views.append(el)

            if not views:
                print("Select one or more views in Revit.")
                return

            # Clear current list, then add selected views
            try:
                vm.DuplicatedViews.Clear()
            except:
                pass

            for v in views:
                try:
                    tmpl = "none"
                    vtid = getattr(v, 'ViewTemplateId', None)
                    if vtid and isinstance(vtid, ElementId) and vtid.IntegerValue > 0:
                        tv = doc.GetElement(vtid)
                        tmpl = tv.Name if tv else "none"
                except:
                    tmpl = "none"

                # add row
                try:
                    vm.add_duplicated_view(v.Id, v.Name, tmpl)
                    # populate per-row template choices
                    pairs = self._collect_templates_for_view(doc, v)
                    try:
                        vm.set_view_templates(v.Id, pairs)
                    except:
                        pass
                except:
                    pass
        except Exception as ex:
            print("Error in _load_selected_views: {}".format(str(ex)))
