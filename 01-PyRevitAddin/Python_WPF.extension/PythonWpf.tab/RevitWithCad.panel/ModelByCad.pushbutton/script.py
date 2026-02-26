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
    Đọc Levels, Categories và FamilySymbol names từ Revit để điền dropdown.
    Chỉ chạy khi Revit API khả dụng.
    Tương thích Revit 2024: dùng FamilyName property + BuiltInParameter fallback
    thay vì truy cập trực tiếp .Name / .Family.Name để tránh crash.
    """
    if not _REVIT_AVAILABLE:
        return
    try:
        doc = __revit__.ActiveUIDocument.Document  # noqa: F821  (pyRevit global)

        # --- Levels ---
        level_elems = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
        def _lv_elev(lv):
            try:
                return lv.Elevation
            except Exception:
                return 0.0
        level_names = []
        for lv in sorted(level_elems, key=_lv_elev):
            try:
                level_names.append(lv.Name)
            except Exception:
                pass
        vm.set_revit_levels(level_names)

        # --- FamilySymbol names + category map ---
        # Revit 2024: KHÔNG gọi s.IsValidObject – có thể gây AccessViolationException.
        symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements())

        cat_type_map = {}
        for s in symbols:
            try:
                cat = s.Category
                if cat is None:
                    continue
                cat_name = cat.Name
                if not cat_name:
                    continue

                # Family name – ưu tiên FamilyName shortcut
                fam_name = ''
                try:
                    fam_name = s.FamilyName or ''
                except Exception:
                    pass
                if not fam_name:
                    try:
                        if s.Family is not None:
                            fam_name = s.Family.Name or ''
                    except Exception:
                        pass
                if not fam_name:
                    continue

                # Type name – fallback sang BuiltInParameter nếu .Name rỗng
                sym_name = ''
                try:
                    sym_name = s.Name or ''
                except Exception:
                    pass
                if not sym_name:
                    try:
                        p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if p is not None:
                            sym_name = p.AsString() or ''
                    except Exception:
                        pass
                if not sym_name:
                    try:
                        p = s.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                        if p is not None:
                            sym_name = p.AsString() or ''
                    except Exception:
                        pass
                if not sym_name:
                    continue

                cat_type_map.setdefault(cat_name, set()).add(
                    "{} : {}".format(fam_name, sym_name)
                )
            except Exception:
                continue

        # Sắp xếp từng nhóm
        cat_type_map_sorted = {k: sorted(v) for k, v in sorted(cat_type_map.items())}
        vm.set_category_type_map(cat_type_map_sorted)

        # Danh sách đầy đủ (dùng khi chưa chọn category)
        type_names = sorted(set(t for types in cat_type_map_sorted.values() for t in types))
        vm.set_revit_family_types(type_names)

        # Categories
        vm.set_revit_categories(sorted(cat_type_map_sorted.keys()))

    except Exception as ex:
        print("_load_revit_data error: {}".format(ex))


class ModelByCadWindow(forms.WPFWindow):
    """Cửa sổ modeless – load XAML, gắn ViewModel + Handler + ExternalEvents."""
    def __init__(self, xaml_file, view_model):
        try:
            forms.WPFWindow.__init__(self, xaml_file)
            self.DataContext = view_model

            # Tag cho DataGrid để handler có thể tìm ra window qua sender.Tag
            self.DgGroups.Tag = self

            # Tạo ExternalEvent để tương tác Revit từ cửa sổ modeless
            if _REVIT_AVAILABLE:
                select_grid_handler = Handler.SelectRevitGridHandler(view_model)
                create_model_handler = Handler.CreateModelHandler(view_model)
                self._ext_select_grid  = ExternalEvent.Create(select_grid_handler)
                self._ext_create_model = ExternalEvent.Create(create_model_handler)

                # Điền dropdown ngay khi mở window
                _load_revit_data(view_model)

            Handler.bind_handlers(self)
        except Exception as ex:
            print('ModelByCadWindow init error: ' + str(ex))
            raise


# Entry point
if __name__ == '__main__':
    vm = ModelByCadViewModel()
    window = ModelByCadWindow('UI.xaml', vm)
    # Show() thay vì ShowDialog() → modeless, người dùng vẫn thao tác được Revit
    window.Show()

