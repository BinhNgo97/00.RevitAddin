# -*- coding: utf-8 -*-
# Copy FamilyTypes / ElementTypes from another RVT into current document

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementType, BuiltInCategory,
    ElementTransformUtils, CopyPasteOptions, Transaction, Transform, Category, ElementId,
    ModelPathUtils, OpenOptions, DetachFromCentralOption  # + added
)
from Autodesk.Revit.UI import IExternalEventHandler
from System import Enum

class ExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        # Inputs
        self.source_file = None         # full path to RVT
        self.category_bic = None        # int (BuiltInCategory)
        self.selected_unique_ids = []   # list[str] of ElementType.UniqueId
        self.action = None              # 'load_categories' | 'load_types' | 'transfer'
        # Outputs
        self.categories = []            # list[{'Name','BIC'}]
        self.types = []                 # list[{'Name','UniqueId'}]
        self.message = ""               # result summary
        # UI busy status
        self.is_busy = False

    def GetName(self):
        return "Copy Types From RVT"

    def Execute(self, app):
        self.is_busy = True
        try:
            if self.action == "load_categories":
                self.categories = self._load_categories(app)
            elif self.action == "load_types":
                self.types = self._load_types(app)
            elif self.action == "transfer":
                self.message = self._transfer_types(app)
        finally:
            self.action = None
            self.is_busy = False

    # ---------- helpers ----------
    def _open_src(self, app):
        """Robust open that works for central/workshared files (detached)."""
        try:
            mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(self.source_file)
            opts = OpenOptions()
            opts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            return app.Application.OpenDocumentFile(mp, opts)
        except:
            # fallback for non-workshared files
            try:
                return app.Application.OpenDocumentFile(self.source_file)
            except:
                return None

    def _load_categories(self, app):
        doc_src = self._open_src(app)
        if not doc_src:
            return []
        try:
            cats = {}
            for et in FilteredElementCollector(doc_src).OfClass(ElementType):
                cat = et.Category
                if not cat:
                    continue
                try:
                    # Use the raw negative int id as BuiltInCategory code
                    bic_int = int(cat.Id.IntegerValue)
                    cats[bic_int] = cat.Name
                except:
                    pass
            items = [{'Name': v, 'BIC': k} for k, v in cats.items()]
            items.sort(key=lambda x: x['Name'])
            return items
        finally:
            try: doc_src.Close(False)
            except: pass

    def _display_name(self, et):
        try:
            fname = getattr(et, "FamilyName", None)
            tname = getattr(et, "Name", None)
            if fname and tname and fname != tname:
                return "{} : {}".format(fname, tname)
            return tname or fname or "Unnamed"
        except:
            return "Unnamed"

    def _load_types(self, app):
        """Return all ElementTypes in the selected category."""
        if self.category_bic is None:
            return []
        doc_src = self._open_src(app)
        if not doc_src:
            return []
        try:
            bic_int = int(self.category_bic)
            result = []
            for et in FilteredElementCollector(doc_src).OfClass(ElementType):
                try:
                    cat = et.Category
                    if cat and int(cat.Id.IntegerValue) == bic_int:
                        result.append({'Name': self._display_name(et), 'UniqueId': et.UniqueId})
                except:
                    pass
            result.sort(key=lambda x: x['Name'])
            return result
        finally:
            try: doc_src.Close(False)
            except: pass

    def _transfer_types(self, app):
        if not self.selected_unique_ids:
            return "No types selected."
        doc_dst = app.ActiveUIDocument.Document
        if not doc_dst:
            return "No active document."

        doc_src = self._open_src(app)
        if not doc_src:
            return "Cannot open source file."

        from Main_Handler import DuplicatesHandler
        copied = 0
        t = None
        try:
            src_ids = []
            for uid in self.selected_unique_ids:
                try:
                    el = doc_src.GetElement(uid)
                    if el: src_ids.append(el.Id)
                except:
                    pass
            if not src_ids:
                return "Selected types not found in source."

            t = Transaction(doc_dst, "Copy Types From File")
            t.Start()
            opts = CopyPasteOptions()
            try:
                opts.SetDuplicateTypeNamesHandler(DuplicatesHandler())
            except:
                pass
            ElementTransformUtils.CopyElements(doc_src, src_ids, doc_dst, Transform.Identity, opts)
            t.Commit()
            copied = len(src_ids)
        except:
            try:
                if t and t.HasStarted(): t.RollBack()
            except:
                pass
        finally:
            try: doc_src.Close(False)
            except: pass

        return "Copied {} type(s).".format(copied)
