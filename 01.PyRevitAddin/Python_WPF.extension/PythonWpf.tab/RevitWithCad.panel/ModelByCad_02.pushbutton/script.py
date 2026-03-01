# -*- coding: utf-8 -*-
import sys
import os
import clr
clr.AddReference('System')
clr.AddReference('System.Data')

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from pyrevit import forms
from ViewModel import ModelByCadViewModel
import Handler

# Revit ExternalEvent API
try:
    from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
    from Autodesk.Revit.DB import (
        FilteredElementCollector, Level, FamilySymbol,
        BuiltInCategory, BuiltInParameter
    )
    _REVIT_AVAILABLE = True
except Exception:
    _REVIT_AVAILABLE = False


def _load_revit_data(vm):
    """
    Đọc Levels, Categories, FamilySymbols từ Revit → điền dropdowns.
    Tương thích Revit 2024: dùng FamilyName + BuiltInParameter fallback.
    """
    if not _REVIT_AVAILABLE:
        return
    try:
        doc = __revit__.ActiveUIDocument.Document  # noqa: F821

        # --- Levels ---
        level_elems = sorted(
            FilteredElementCollector(doc).OfClass(Level).ToElements(),
            key=lambda lv: lv.Elevation
        )
        level_names = []
        for lv in level_elems:
            try: level_names.append(lv.Name)
            except Exception: pass
        vm.set_revit_levels(level_names)

        # --- FamilySymbols + Category map ---
        symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements())
        cat_type_map = {}

        for s in symbols:
            try:
                cat = s.Category
                if cat is None: continue
                cat_name = cat.Name
                if not cat_name: continue

                fam_name = ''
                try: fam_name = s.FamilyName or ''
                except Exception: pass
                if not fam_name:
                    try:
                        if s.Family: fam_name = s.Family.Name or ''
                    except Exception: pass
                if not fam_name: continue

                sym_name = ''
                try: sym_name = s.Name or ''
                except Exception: pass
                if not sym_name:
                    try:
                        p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if p: sym_name = p.AsString() or ''
                    except Exception: pass
                if not sym_name:
                    try:
                        p = s.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                        if p: sym_name = p.AsString() or ''
                    except Exception: pass
                if not sym_name: continue

                cat_type_map.setdefault(cat_name, set()).add(
                    "{} : {}".format(fam_name, sym_name)
                )
            except Exception:
                continue

        cat_type_map_sorted = {k: sorted(v) for k, v in sorted(cat_type_map.items())}
        vm.set_category_type_map(cat_type_map_sorted)

        type_names = sorted(set(t for types in cat_type_map_sorted.values() for t in types))
        vm.set_revit_family_types(type_names)
        vm.set_revit_categories(sorted(cat_type_map_sorted.keys()))

    except Exception as ex:
        print("_load_revit_data error: {}".format(ex))


class ModelByCadWindow(forms.WPFWindow):
    """Cửa sổ modeless – load XAML, gắn ViewModel + ExternalEvents + Handlers."""

    def __init__(self, xaml_file, view_model):
        try:
            forms.WPFWindow.__init__(self, xaml_file)
            self.DataContext = view_model

            # Tag DataGrids (handler cần tìm window qua sender.Tag)
            self.DgGroups.Tag      = self
            self.DgConditions.Tag  = self
            self.DgRules.Tag       = self

            if _REVIT_AVAILABLE:
                # ExternalEvent: chọn đường tham chiếu Revit per-condition
                select_line_handler      = Handler.SelectRevitLineHandler(view_model)
                self._ext_select_line    = select_line_handler
                self._ext_select_line_event = ExternalEvent.Create(select_line_handler)

                # ExternalEvent: chọn đường tham chiếu Revit global
                select_grid_handler    = Handler.SelectRevitGridHandler(view_model)
                self._ext_select_grid  = ExternalEvent.Create(select_grid_handler)

                # ExternalEvent: Create Model (all conditions – legacy)
                create_model_handler   = Handler.CreateModelHandler(view_model)
                self._ext_create_model = ExternalEvent.Create(create_model_handler)

                # ExternalEvent: Create Model for a single condition
                create_single_handler                = Handler.CreateModelSingleHandler(view_model)
                create_single_handler._window        = self
                self._ext_create_model_single        = create_single_handler
                self._ext_create_model_single_event  = ExternalEvent.Create(create_single_handler)

                # Điền data Revit vào dropdowns
                _load_revit_data(view_model)

            # Gắn tất cả event handlers
            Handler.bind_handlers(self)

        except Exception as ex:
            print('ModelByCadWindow init error: ' + str(ex))
            raise


# ── Entry point ────────────────────────────────────────────
if __name__ == '__main__':
    vm     = ModelByCadViewModel()
    window = ModelByCadWindow('UI.xaml', vm)
    window.Show()   # modeless → người dùng vẫn thao tác được Revit
