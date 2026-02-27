# -*- coding: utf-8 -*-
"""
ViewModel.py - ViewModel cho tool Link Multiple CAD.
"""
import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Collections.ObjectModel import ObservableCollection
from aGeneral.ViewModel_Base import ViewModel_BaseEventHandler


# =============================================
# DwgEntry – item trong dropdown cột File Cad
# =============================================
class DwgEntry(object):
    """Wrapper đơn giản: lưu full path + tên file để hiển thị trong ComboBox."""
    def __init__(self, full_path):
        self._path = full_path

    @property
    def Name(self):
        return os.path.basename(self._path)

    @property
    def Path(self):
        return self._path

    def __repr__(self):
        return self.Name


# =============================================
# CadFileRow – 1 hàng trong DataGrid
# =============================================
class CadFileRow(ViewModel_BaseEventHandler):
    """
    Đại diện cho 1 file DWG cần link vào Revit.
    Binding vào DataGrid: SelectedEntry (ComboBox File Cad) | BaseLevel (ComboBox)
    """
    def __init__(self):
        ViewModel_BaseEventHandler.__init__(self)
        self._entry      = None   # DwgEntry đang được chọn
        self._base_level = ''

    # ---- SelectedEntry – item người dùng chọn trong combo cột File Cad ----
    @property
    def SelectedEntry(self):
        return self._entry

    @SelectedEntry.setter
    def SelectedEntry(self, value):
        self._entry = value
        self.OnPropertyChanged('SelectedEntry')
        self.OnPropertyChanged('FileName')

    # ---- FileName – computed, hiển thị trong CellTemplate ----
    @property
    def FileName(self):
        return self._entry.Name if self._entry else u'-- chọn file --'

    # ---- FilePath – để LinkCadHandler đọc ----
    @property
    def FilePath(self):
        return self._entry.Path if self._entry else ''

    # ---- BaseLevel ----
    @property
    def BaseLevel(self):
        return self._base_level

    @BaseLevel.setter
    def BaseLevel(self, value):
        self._base_level = value or ''
        self.OnPropertyChanged('BaseLevel')

    def is_ready(self):
        """Kiểm tra đã chọn file và level (đủ điều kiện để link)."""
        return bool(self._entry and self._base_level)


# =============================================
# LinkMultiCadViewModel – ViewModel chính
# =============================================
class LinkMultiCadViewModel(ViewModel_BaseEventHandler):
    """
    ViewModel chính cho tool Link Multiple CAD.
    """
    def __init__(self):
        ViewModel_BaseEventHandler.__init__(self)

        # Danh sách hàng trong DataGrid
        self._cad_files = ObservableCollection[CadFileRow]()

        # Danh sách DwgEntry để điền vào ComboBox cột File Cad
        self._available_dwg = ObservableCollection[object]()

        # Dropdown Levels từ Revit
        self._levels = ObservableCollection[object]()

        # Layer chứa đường baseline trong mỗi file CAD
        self._layer_origin_name = ''

        # Folder đang được chọn
        self.folder_path = ''

        # Đường tham chiếu đã chọn trong Revit (lưu bởi SelectBaseLineHandler)
        # Kiểu: Revit Curve element (Grid / ModelLine / DetailLine...)
        self.revit_baseline = None

        # Hàng đang được chọn trong DataGrid
        self._selected_file = None

        # Status text
        self._status = 'Sẵn sàng.'

    # ---- CadFiles ----
    @property
    def CadFiles(self):
        return self._cad_files

    # ---- AvailableDwgFiles – items cho ComboBox cột File Cad ----
    @property
    def AvailableDwgFiles(self):
        return self._available_dwg

    # ---- Levels (dropdown) ----
    @property
    def Levels(self):
        return self._levels

    def set_revit_levels(self, names):
        """Cập nhật danh sách level (gọi từ script.py sau khi đọc Revit)."""
        self._levels.Clear()
        for n in names:
            self._levels.Add(n)

    # ---- LayerOriginName ----
    @property
    def LayerOriginName(self):
        return self._layer_origin_name

    @LayerOriginName.setter
    def LayerOriginName(self, value):
        self._layer_origin_name = value or ''
        self.OnPropertyChanged('LayerOriginName')

    # ---- SelectedFile ----
    @property
    def SelectedFile(self):
        return self._selected_file

    @SelectedFile.setter
    def SelectedFile(self, value):
        self._selected_file = value
        self.OnPropertyChanged('SelectedFile')

    # ---- Status ----
    @property
    def Status(self):
        return self._status

    @Status.setter
    def Status(self, value):
        self._status = value or ''
        self.OnPropertyChanged('Status')

    # ---- Helpers: quét folder / thêm / xóa dòng ----
    def scan_folder(self, folder_path):
        """
        Quét thư mục, cập nhật AvailableDwgFiles (dùng cho ComboBox).
        Trả về số file .dwg tìm thấy.
        """
        if not folder_path or not os.path.isdir(folder_path):
            return 0
        self.folder_path = folder_path
        self._available_dwg.Clear()
        count = 0
        for fname in sorted(os.listdir(folder_path)):
            if fname.lower().endswith('.dwg'):
                full = os.path.join(folder_path, fname)
                self._available_dwg.Add(DwgEntry(full))
                count += 1
        self.OnPropertyChanged('AvailableDwgFiles')
        return count

    def add_empty_row(self):
        """Thêm 1 hàng trống vào DataGrid để người dùng chọn file qua ComboBox."""
        self._cad_files.Add(CadFileRow())

    def remove_rows(self, rows):
        """
        Xóa danh sách hàng (hỗ trợ multi-select).
        rows: iterable của CadFileRow.
        """
        for row in list(rows):
            try:
                self._cad_files.Remove(row)
            except Exception:
                pass
        self._selected_file = None
        self.OnPropertyChanged('SelectedFile')

    def has_baseline(self):
        return self.revit_baseline is not None
