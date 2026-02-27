# -*- coding: utf-8 -*-
"""
script.py - Entry point cho tool Link Multiple CAD (pyRevit pushbutton).

Workflow:
  1. Đọc danh sách Level từ Revit.
  2. Tạo ViewModel + WPF window.
  3. Tạo ExternalEvents (SelectBaseLine, LinkCad).
  4. Hiển thị window ở chế độ modeless (Show).
"""
import sys
import os
import clr
clr.AddReference('System')

# Thêm thư mục extension vào sys.path để import aGeneral
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from pyrevit import forms
from ViewModel import LinkMultiCadViewModel
import Handler

# ============================================================
#   Revit API (optional – chỉ khả dụng khi chạy trong Revit)
# ============================================================
try:
    from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
    from Autodesk.Revit.DB import FilteredElementCollector, Level
    _REVIT_AVAILABLE = True
except Exception:
    _REVIT_AVAILABLE = False


# ============================================================
#   Đọc dữ liệu Revit vào ViewModel
# ============================================================

def _load_revit_levels(vm):
    """
    Đọc tất cả Level từ tài liệu Revit hiện tại và điền vào vm.Levels.
    Sắp xếp theo elevation tăng dần.
    """
    if not _REVIT_AVAILABLE:
        return
    try:
        doc = __revit__.ActiveUIDocument.Document  # noqa: F821 – pyRevit global

        level_elems = list(
            FilteredElementCollector(doc).OfClass(Level).ToElements()
        )

        def _elev(lv):
            try:
                return lv.Elevation
            except Exception:
                return 0.0

        names = []
        for lv in sorted(level_elems, key=_elev):
            try:
                names.append(lv.Name)
            except Exception:
                pass

        vm.set_revit_levels(names)

    except Exception as ex:
        print("_load_revit_levels error: {}".format(ex))


# ============================================================
#   WPF Window
# ============================================================

class LinkMultiCadWindow(forms.WPFWindow):
    """
    Cửa sổ modeless cho tool Link Multiple CAD.
    Load XAML, gắn ViewModel + Handler + ExternalEvents.
    """

    def __init__(self, xaml_file, view_model):
        try:
            forms.WPFWindow.__init__(self, xaml_file)
            self.DataContext = view_model

            # Tạo ExternalEvents (chỉ khi Revit API khả dụng)
            if _REVIT_AVAILABLE:
                select_baseline_handler = Handler.SelectBaseLineHandler(view_model)
                link_cad_handler        = Handler.LinkCadHandler(view_model)

                self._ext_select_baseline = ExternalEvent.Create(select_baseline_handler)
                self._ext_link_cad        = ExternalEvent.Create(link_cad_handler)

                # Điền Levels ngay khi mở window
                _load_revit_levels(view_model)

            # Gắn tất cả button handlers
            Handler.bind_handlers(self)

            view_model.Status = u"Sẵn sàng. Chọn folder CAD và nhập Layer Origin Name."

        except Exception as ex:
            print("LinkMultiCadWindow init error: {}".format(ex))
            raise


# ============================================================
#   Entry point (pyRevit chạy __file__)
# ============================================================
if __name__ == '__main__':
    vm     = LinkMultiCadViewModel()
    window = LinkMultiCadWindow('UI.xaml', vm)
    # Show() → modeless: người dùng vẫn thao tác được trong Revit
    window.Show()
