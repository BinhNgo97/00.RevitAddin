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
    Đọc Levels, Categories (BuiltInCategory) và FamilySymbol names từ Revit
    để điền dropdown trong DataGrid.
    Chỉ chạy khi Revit API khả dụng.
    """
    if not _REVIT_AVAILABLE:
        print("[DEBUG] _load_revit_data: Revit API không khả dụng, bỏ qua.")
        return
    try:
        doc = __revit__.ActiveUIDocument.Document  # noqa: F821  (pyRevit global)
        print("[DEBUG] _load_revit_data: Bắt đầu đọc dữ liệu Revit...")

        # --- Levels ---
        # Chuyển sang list Python trước khi sort để tránh lỗi lazy-collection
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
        print("[DEBUG] Levels tìm thấy: {}".format(level_names))

        # --- FamilySymbol names + category map ---
        # Revit 2024: KHÔNG gọi s.IsValidObject – property đó có thể gây
        # AccessViolationException (hard crash) với các element native bị invalid.
        # Dùng thuần try/except cho từng thuộc tính.
        symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements())
        print("[DEBUG] Tổng FamilySymbol tìm thấy: {}".format(len(symbols)))

        cat_type_map = {}
        error_count = 0
        skip_no_cat = 0
        skip_no_catname = 0
        skip_no_famname = 0
        skip_no_symname = 0
        ok_count = 0
        for i, s in enumerate(symbols):
            try:
                cat = s.Category
                if cat is None:
                    skip_no_cat += 1
                    continue
                cat_name = cat.Name
                if not cat_name:
                    skip_no_catname += 1
                    continue

                # Ưu tiên s.FamilyName (shortcut an toàn hơn s.Family.Name)
                fam_name = ''
                try:
                    fam_name = s.FamilyName or ''
                except Exception:
                    pass
                # Fallback: thử s.Family.Name nếu FamilyName rỗng
                if not fam_name:
                    try:
                        if s.Family is not None:
                            fam_name = s.Family.Name or ''
                    except Exception:
                        pass
                if not fam_name:
                    skip_no_famname += 1
                    if skip_no_famname <= 2:
                        # In vài mẫu để debug
                        try:
                            print("[DEBUG] skip_no_famname cat={} sym={}".format(cat_name, s.Name))
                        except Exception:
                            print("[DEBUG] skip_no_famname cat={} (sym.Name loi)".format(cat_name))
                    continue

                sym_name = ''
                sym_ex2 = None
                try:
                    sym_name = s.Name or ''
                except Exception as _e:
                    sym_ex2 = _e
                # Fallback 1: BuiltInParameter.SYMBOL_NAME_PARAM
                if not sym_name:
                    try:
                        p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if p is not None:
                            sym_name = p.AsString() or ''
                    except Exception:
                        pass
                # Fallback 2: ALL_MODEL_TYPE_NAME
                if not sym_name:
                    try:
                        p = s.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                        if p is not None:
                            sym_name = p.AsString() or ''
                    except Exception:
                        pass
                # Debug 3 symbol đầu tiên
                if ok_count + skip_no_symname < 3:
                    print("[DEBUG] s.Name={!r}  sym_ex={} fam={}  sym_final={!r}".format(
                        sym_name, sym_ex2, fam_name, sym_name))
                if not sym_name:
                    skip_no_symname += 1
                    continue

                type_name = "{} : {}".format(fam_name, sym_name)
                cat_type_map.setdefault(cat_name, set()).add(type_name)
                ok_count += 1
            except Exception as sym_ex:
                error_count += 1
                if error_count <= 3:
                    print("[DEBUG] Symbol lỗi: {}".format(sym_ex))
                continue

        print("[DEBUG] Symbol stats: total={} ok={} err={} no_cat={} no_catname={} no_famname={} no_symname={}".format(
            len(symbols), ok_count, error_count, skip_no_cat, skip_no_catname, skip_no_famname, skip_no_symname))

        # Sắp xếp từng nhóm
        cat_type_map_sorted = {k: sorted(v) for k, v in sorted(cat_type_map.items())}
        vm.set_category_type_map(cat_type_map_sorted)

        # Danh sách đầy đủ (dùng khi chưa chọn category)
        type_names = sorted(set(t for types in cat_type_map_sorted.values() for t in types))
        vm.set_revit_family_types(type_names)

        # --- Categories – lấy từ map keys ---
        cat_list = sorted(cat_type_map_sorted.keys())
        vm.set_revit_categories(cat_list)
        print("[DEBUG] _load_revit_data XONG. Categories đã set: {}".format(len(cat_list)))
        print("[DEBUG] vm.Categories count: {}".format(vm.Categories.Count))

    except Exception as ex:
        import traceback
        print("_load_revit_data error: {}\n{}".format(ex, traceback.format_exc()))


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

