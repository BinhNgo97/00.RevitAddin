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
        BuiltInCategory
    )
    _REVIT_AVAILABLE = True
except Exception:
    _REVIT_AVAILABLE = False


def _load_revit_data(vm):
    """
    Đọc Levels, Categories (BuiltInCategory) và FamilySymbol names từ Revit
    để điền dropdown trong DataGrid.
    Chỉ chạy khi Revit API khả dụng.
    """
    if not _REVIT_AVAILABLE:
        return
    try:
        doc = __revit__.ActiveUIDocument.Document  # noqa: F821  (pyRevit global)

        # --- Levels ---
        levels = sorted(
            FilteredElementCollector(doc).OfClass(Level).ToElements(),
            key=lambda l: l.Elevation
        )
        vm.set_revit_levels([lv.Name for lv in levels])

        # --- FamilySymbol names + category map ---
        symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()

        # Map: category_name → sorted list of "FamilyName : TypeName"
        cat_type_map = {}
        for s in symbols:
            if s.Family is None or s.Category is None:
                continue
            cat_name  = s.Category.Name
            type_name = "{} : {}".format(s.Family.Name, s.Name)
            cat_type_map.setdefault(cat_name, set()).add(type_name)
        # Sắp xếp từng nhóm
        cat_type_map_sorted = {k: sorted(v) for k, v in sorted(cat_type_map.items())}
        vm.set_category_type_map(cat_type_map_sorted)

        # Danh sách đầy đủ (dùng khi chưa chọn category)
        type_names = sorted(set(t for types in cat_type_map_sorted.values() for t in types))
        vm.set_revit_family_types(type_names)

        # --- Categories – lấy từ map keys ---
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

