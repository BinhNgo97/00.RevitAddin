# -*- coding: utf-8 -*-
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Collections.ObjectModel import ObservableCollection
from aGeneral.ViewModel_Base import ViewModel_BaseEventHandler


# =============================================
# CadGroup – 1 row trong DataGrid "Set up data type"
# =============================================
class CadGroup(ViewModel_BaseEventHandler):
    """
    Đại diện cho 1 nhóm elements có cùng kích thước trong CAD.
    Binding vào DataGrid: Label | Category | FamilyType | BaseLevel
    """
    def __init__(self, label, shape, elements, w=0, h=0, dia=0, vm=None, source_type='point'):
        ViewModel_BaseEventHandler.__init__(self)
        self.label    = label       # VD: "REC: 200x300", "CIR: 400", "BEA: 200x400"
        self.shape    = shape       # 'REC' | 'CIR' | 'BEA'
        self.elements = elements    # list[CadElement / _MergedPolyline / CadBeamPair]
        self.w        = w
        self.h        = h
        self.dia      = dia
        self._vm      = vm
        # 'point' = cột/móng đặt theo XYZ tâm
        # 'line'  = dầm đặt theo Curve (start–end)
        self.source_type = source_type

        self._category    = ''
        self._family_type = ''
        self._base_level  = ''

    # ---- Label (read-only) ----
    @property
    def Label(self):
        return self.label

    # ---- Category ----
    @property
    def Category(self):
        return self._category

    @Category.setter
    def Category(self, value):
        self._category = value
        self.OnPropertyChanged('Category')
        # Khi đổi category → reset family type + cập nhật danh sách lọc
        self._family_type = ''
        self.OnPropertyChanged('FamilyType')
        self.OnPropertyChanged('FilteredFamilyTypes')

    # ---- FamilyType ----
    @property
    def FamilyType(self):
        return self._family_type

    @FamilyType.setter
    def FamilyType(self, value):
        self._family_type = value
        self.OnPropertyChanged('FamilyType')

    # ---- BaseLevel ----
    @property
    def BaseLevel(self):
        return self._base_level

    @BaseLevel.setter
    def BaseLevel(self, value):
        self._base_level = value
        self.OnPropertyChanged('BaseLevel')

    # ---- FilteredFamilyTypes (read-only, phụ thuộc vào Category) ----
    @property
    def FilteredFamilyTypes(self):
        """Trả về danh sách Family Type thuộc Category hiện tại."""
        if self._vm is not None and self._category:
            return self._vm.get_types_for_category(self._category)
        return []

    def is_beam(self):
        return self.source_type == 'line'

    def is_ready(self):
        """Kiểm tra đã điền đủ thông tin để tạo model chưa."""
        return bool(self._category and self._family_type and self._base_level)

    # ---- Count (số lượng element trong nhóm, read-only) ----
    @property
    def Count(self):
        return len(self.elements)

    def notify_count_changed(self):
        self.OnPropertyChanged('Count')


# =============================================
# Main ViewModel
# =============================================
class ModelByCadViewModel(ViewModel_BaseEventHandler):
    """
    ViewModel chính cho tool Model By CAD.
    """
    def __init__(self):
        ViewModel_BaseEventHandler.__init__(self)

        # Dữ liệu CAD
        self.cad_elements      = []   # list[CadElement | _MergedPolyline] – cột/móng
        self.beam_elements     = []   # list[CadBeamPair] – dầm
        self.cad_grid_elements = []   # list[CadElement] – 1 đường tham chiếu CAD
        self.revit_grids       = []   # list[Revit Grid / Line]

        # Danh sách nhóm (binding DataGrid bên phải)
        self._cad_groups = ObservableCollection[CadGroup]()

        # Danh sách dropdown lấy từ Revit
        self._categories   = ObservableCollection[object]()
        self._family_types = ObservableCollection[object]()
        self._levels       = ObservableCollection[object]()

        # Map: category_name → list of "Family : Type" strings
        self._category_type_map = {}

        # Nhóm đang được chọn trong DataGrid
        self._selected_group = None

        # Element đơn lẻ đang được chọn trong Preview (canvas click)
        self._selected_element = None

        self._status = ''

    # ---- CadGroups (DataGrid source) ----
    @property
    def CadGroups(self):
        return self._cad_groups

    def rebuild_groups(self, groups_data):
        """
        Cập nhật CadGroups từ kết quả group_elements_by_label() (cột/móng).
        Giữ nguyên các hàng dầm (source_type='line') đã có.
        """
        beam_rows = [g for g in self._cad_groups if g.is_beam()]
        self._cad_groups.Clear()
        for g in groups_data:
            item = CadGroup(
                label       = g['label'],
                shape       = g['shape'],
                elements    = g['elements'],
                w           = g.get('w', 0),
                h           = g.get('h', 0),
                dia         = g.get('dia', 0),
                vm          = self,
                source_type = 'point',
            )
            self._cad_groups.Add(item)
        # Thêm lại các hàng dầm cũ (nếu đã chọn trước)
        for br in beam_rows:
            self._cad_groups.Add(br)

    def rebuild_beam_groups(self, groups_data):
        """
        Cập nhật / ghi đè các hàng dầm trong CadGroups.
        Giữ nguyên các hàng cột/móng (source_type='point').
        """
        point_rows = [g for g in self._cad_groups if not g.is_beam()]
        self._cad_groups.Clear()
        for pr in point_rows:
            self._cad_groups.Add(pr)
        for g in groups_data:
            item = CadGroup(
                label       = g['label'],
                shape       = g['shape'],
                elements    = g['elements'],
                w           = g.get('w', 0),
                h           = g.get('h', 0),
                vm          = self,
                source_type = 'line',
            )
            self._cad_groups.Add(item)

    def merge_into_groups(self, groups_data):
        """
        Bổ sung các elements mới vào nhóm đã có (theo label).
        - Nếu label đã tồn tại → append elements, cập nhật Count.
        - Nếu chưa có   → tạo nhóm mới (giữ nguyên dầm).
        """
        existing = {g.label: g for g in self._cad_groups if not g.is_beam()}
        for g in groups_data:
            lbl = g['label']
            if lbl in existing:
                for elem in g['elements']:
                    existing[lbl].elements.append(elem)
                existing[lbl].notify_count_changed()
            else:
                item = CadGroup(
                    label       = lbl,
                    shape       = g['shape'],
                    elements    = list(g['elements']),
                    w           = g.get('w', 0),
                    h           = g.get('h', 0),
                    dia         = g.get('dia', 0),
                    vm          = self,
                    source_type = 'point',
                )
                # Chèn trước các hàng dầm (bảo toàn thứ tự: cột/móng rồi dầm)
                beam_start = next(
                    (i for i, gr in enumerate(self._cad_groups) if gr.is_beam()),
                    len(self._cad_groups)
                )
                self._cad_groups.Insert(beam_start, item)

    def merge_into_beam_groups(self, groups_data):
        """
        Bổ sung các dầm mới vào nhóm dầm đã có (theo label).
        - Nếu label đã tồn tại → append elements, cập nhật Count.
        - Nếu chưa có   → tạo nhóm dầm mới (giữ nguyên cột/móng).
        """
        existing = {g.label: g for g in self._cad_groups if g.is_beam()}
        for g in groups_data:
            lbl = g['label']
            if lbl in existing:
                for pair in g['elements']:
                    existing[lbl].elements.append(pair)
                existing[lbl].notify_count_changed()
            else:
                item = CadGroup(
                    label       = lbl,
                    shape       = g['shape'],
                    elements    = list(g['elements']),
                    w           = g.get('w', 0),
                    h           = g.get('h', 0),
                    vm          = self,
                    source_type = 'line',
                )
                self._cad_groups.Add(item)

    # ---- Revit dropdown lists ----
    @property
    def Categories(self):
        return self._categories

    @property
    def FamilyTypes(self):
        return self._family_types

    @property
    def Levels(self):
        return self._levels

    def set_revit_categories(self, names):
        self._categories.Clear()
        for n in names:
            self._categories.Add(n)

    def set_category_type_map(self, mapping):
        """mapping: dict {cat_name: [type_string, ...]}"""
        self._category_type_map = mapping

    def get_types_for_category(self, category_name):
        """Trả về list type strings cho category đã cho."""
        return self._category_type_map.get(category_name, [])

    def set_revit_family_types(self, names):
        self._family_types.Clear()
        for n in names:
            self._family_types.Add(n)

    def set_revit_levels(self, names):
        self._levels.Clear()
        for n in names:
            self._levels.Add(n)

    # ---- Selected group ----
    @property
    def SelectedGroup(self):
        return self._selected_group

    @SelectedGroup.setter
    def SelectedGroup(self, value):
        self._selected_group = value
        self.OnPropertyChanged('SelectedGroup')

    # ---- SelectedElement (chọn từng element trong canvas) ----
    @property
    def SelectedElement(self):
        return self._selected_element

    @SelectedElement.setter
    def SelectedElement(self, value):
        self._selected_element = value
        self.OnPropertyChanged('SelectedElement')

    # ---- Helpers xóa element / nhóm ----
    def remove_element(self, elem):
        """
        Xóa 1 element khỏi nhóm chứa nó.
        Nếu nhóm rống sau khi xóa thì xóa luôn nhóm.
        Trả về True nếu thành công.
        """
        for group in list(self._cad_groups):
            if elem in group.elements:
                group.elements.remove(elem)
                if elem in self.cad_elements:
                    self.cad_elements.remove(elem)
                if elem in self.beam_elements:
                    self.beam_elements.remove(elem)
                group.notify_count_changed()
                if not group.elements:
                    self._cad_groups.Remove(group)
                return True
        return False

    def remove_group(self, group):
        """Xóa toàn bộ nhóm và tất cả elements của nó."""
        for elem in list(group.elements):
            if elem in self.cad_elements:
                self.cad_elements.remove(elem)
            if elem in self.beam_elements:
                self.beam_elements.remove(elem)
        group.elements[:] = []
        if group in list(self._cad_groups):
            self._cad_groups.Remove(group)

    def clear_all(self):
        """Xóa toàn bộ dữ liệu CAD + DataGrid (Refresh)."""
        self._cad_groups.Clear()
        self.cad_elements      = []
        self.beam_elements     = []
        self.cad_grid_elements = []
        self._selected_group   = None
        self._selected_element = None
        self.OnPropertyChanged('SelectedGroup')
        self.OnPropertyChanged('SelectedElement')

    # ---- Status ----
    @property
    def Status(self):
        return self._status

    @Status.setter
    def Status(self, value):
        self._status = value
        self.OnPropertyChanged('Status')

    def has_grid_transform(self):
        return len(self.cad_grid_elements) >= 1 and len(self.revit_grids) >= 1
