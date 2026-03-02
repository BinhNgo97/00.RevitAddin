# -*- coding: utf-8 -*-
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Collections.ObjectModel import ObservableCollection
from aGeneral.ViewModel_Base import ViewModel_BaseEventHandler


# =============================================
# COLOR PALETTE – mỗi condition 1 màu riêng trên Canvas
# =============================================
_CONDITION_COLORS = [
    '#00E5FF',  # Cyan
    '#AEEA00',  # Yellow-Green
    '#FF6D00',  # Orange
    '#E040FB',  # Purple
    '#69F0AE',  # Mint Green
    '#FFD740',  # Amber
    '#FF4081',  # Pink
    '#40C4FF',  # Light Blue
    '#F4FF81',  # Lime Yellow
    '#EA80FC',  # Light Purple
]
_color_index = [0]


def _next_color():
    c = _CONDITION_COLORS[_color_index[0] % len(_CONDITION_COLORS)]
    _color_index[0] += 1
    return c


# =============================================
# RuleRow – 1 dòng trong Bảng 2  (Parameter | Ruler | Value)
# =============================================
class RuleRow(ViewModel_BaseEventHandler):
    """
    1 rule filter:
      - Layer Name + Equal         → Value là dropdown layer name
      - Length / Distance + Ruler  → Value là giá trị số
    Logic giữa các dòng trong 1 condition: OR
    """
    PARAMETERS     = ['Layer Name', 'Length', 'Min Beam Distance', 'Max Beam Distance', 'Text Layer']
    RULERS_LAYER   = ['Equal']
    RULERS_NUMERIC = ['is greater than', 'is less than', 'Equal']

    def __init__(self, parameter='Layer Name', ruler='Equal', value=''):
        ViewModel_BaseEventHandler.__init__(self)
        self._parameter = parameter
        self._ruler     = ruler
        self._value     = value

    # ---- Parameter ----
    @property
    def Parameter(self):
        return self._parameter

    @Parameter.setter
    def Parameter(self, value):
        self._parameter = value
        # Auto adjust ruler
        if value in ('Layer Name', 'Text Layer'):
            self._ruler = 'Equal'
        else:
            if self._ruler == 'Equal':
                self._ruler = 'is greater than'
        self._value = ''
        self.OnPropertyChanged('Parameter')
        self.OnPropertyChanged('Ruler')
        self.OnPropertyChanged('Value')
        self.OnPropertyChanged('IsLayerRule')
        self.OnPropertyChanged('AvailableRulers')

    # ---- Ruler ----
    @property
    def Ruler(self):
        return self._ruler

    @Ruler.setter
    def Ruler(self, value):
        self._ruler = value
        self.OnPropertyChanged('Ruler')

    # ---- Value ----
    @property
    def Value(self):
        return self._value

    @Value.setter
    def Value(self, value):
        self._value = value
        self.OnPropertyChanged('Value')

    # ---- Computed ----
    @property
    def AvailableRulers(self):
        if self._parameter == 'Layer Name':
            return self.RULERS_LAYER
        return self.RULERS_NUMERIC

    @property
    def IsLayerRule(self):
        """True → Value ô là dropdown layer; False → TextBox số."""
        return self._parameter in ('Layer Name', 'Text Layer')

    # ---- Serialization ----
    def to_dict(self):
        return {'parameter': self._parameter, 'ruler': self._ruler, 'value': self._value}

    @classmethod
    def from_dict(cls, d):
        return cls(d.get('parameter', 'Layer Name'),
                   d.get('ruler', 'Equal'),
                   d.get('value', ''))


# =============================================
# CadGroup – 1 nhóm kích thước trong Set up data type
# =============================================
class CadGroup(ViewModel_BaseEventHandler):
    """
    Nhóm elements cùng label (shape + kích thước).
    Binding DataGrid Set up data type: Qty | Data From Cad | Family Type
    """
    def __init__(self, label, shape, elements, w=0, h=0, dia=0,
                 vm=None, source_type='point', condition=None):
        ViewModel_BaseEventHandler.__init__(self)
        self.label       = label
        self.shape       = shape
        self.elements    = elements     # list[element]
        self.w           = w
        self.h           = h
        self.dia         = dia
        self._vm         = vm
        self.source_type = source_type  # 'point' = cột/móng | 'line' = dầm
        self.condition   = condition    # ConditionRow cha (để lấy category + color)
        self._family_type = ''

    @property
    def Label(self):
        return self.label

    @property
    def Category(self):
        """Lấy category từ ConditionRow cha."""
        if self.condition is not None:
            return self.condition.Category
        return ''

    @property
    def FamilyType(self):
        return self._family_type

    @FamilyType.setter
    def FamilyType(self, value):
        self._family_type = value
        self.OnPropertyChanged('FamilyType')

    @property
    def FilteredFamilyTypes(self):
        """Danh sách Family Type lọc theo Category (từ Revit)."""
        if self._vm and self.Category:
            return self._vm.get_types_for_category(self.Category)
        return []

    @property
    def Count(self):
        return len(self.elements)

    def notify_count_changed(self):
        self.OnPropertyChanged('Count')

    def is_beam(self):
        return self.source_type == 'line'

    def is_ready(self):
        """Đã chọn Family Type chưa."""
        return bool(self._family_type)


# =============================================
# SelectedBeamInfoRow – thông tin dầm được click trên canvas
# =============================================
class SelectedBeamInfoRow(ViewModel_BaseEventHandler):
    """
    Binding cho floating overlay panel khi user click 1 beam_axis trên canvas.
    Cho phép sửa Location Type và Family Type riêng cho từng element.
    """
    LOCATION_TYPES = [u'Auto', u'Left', u'Right', u'Center', u'Origin']
    _LOC_TO_INT = {u'Auto': None, u'Left': 0, u'Right': 1, u'Center': 2, u'Origin': 3}
    _INT_TO_LOC = {None: u'Auto', 0: u'Left', 1: u'Right', 2: u'Center', 3: u'Origin'}

    def __init__(self, elem, group, condition, vm):
        ViewModel_BaseEventHandler.__init__(self)
        self._elem      = elem       # BeamAxis
        self._group     = group      # CadGroup hiện tại của elem
        self._condition = condition  # ConditionRow chứa elem
        self._vm        = vm         # ModelByCadViewModel

    @property
    def BeamLabel(self):
        lbl = getattr(self._elem, 'text_label', None)
        return lbl if lbl else u'w={}'.format(getattr(self._elem, 'width', 0))

    @property
    def LocationType(self):
        return self._INT_TO_LOC.get(
            getattr(self._elem, 'location_type_override', None), u'Auto')

    @LocationType.setter
    def LocationType(self, v):
        self._elem.location_type_override = self._LOC_TO_INT.get(v, None)
        self.OnPropertyChanged('LocationType')

    @property
    def FamilyType(self):
        override = getattr(self._elem, 'family_type_override', None)
        return override if override is not None else self._group.FamilyType

    @FamilyType.setter
    def FamilyType(self, v):
        if v == self.FamilyType:
            return
        self._elem.family_type_override = v
        self._rebalance_groups(v)
        self.OnPropertyChanged('FamilyType')

    @property
    def FilteredFamilyTypes(self):
        if self._vm and self._condition:
            return self._vm.get_types_for_category(self._condition.Category)
        return []

    def _rebalance_groups(self, new_fam_type):
        """Chuyển elem sang nhóm khác khi FamilyType thay đổi."""
        elem      = self._elem
        old_group = self._group
        condition = self._condition

        # 1) Xóa elem khỏi nhóm cũ
        if elem in old_group.elements:
            old_group.elements.remove(elem)
            old_group.notify_count_changed()

        # 2) Xóa nhóm rỗng
        if len(old_group.elements) == 0:
            if old_group in list(condition.cad_groups):
                condition.cad_groups.remove(old_group)

        # 3) Tìm nhóm đã có cùng FamilyType trong condition
        target_group = None
        for g in condition.cad_groups:
            if g.FamilyType == new_fam_type and g.shape == old_group.shape:
                target_group = g
                break

        if target_group is not None:
            target_group.elements.append(elem)
            target_group.notify_count_changed()
            self._group = target_group
        else:
            # 4) Tạo nhóm mới cho element này
            new_group = CadGroup(
                label       = old_group.label,
                shape       = old_group.shape,
                elements    = [elem],
                w           = old_group.w,
                h           = old_group.h,
                dia         = old_group.dia,
                vm          = self._vm,
                source_type = old_group.source_type,
                condition   = condition,
            )
            new_group.FamilyType = new_fam_type
            condition.cad_groups.append(new_group)
            self._group = new_group

        # 5) Cập nhật PreviewGroups
        if self._vm:
            self._vm.refresh_preview_groups()


# =============================================
# ConditionRow – 1 dòng trong Bảng 3
# =============================================
class ConditionRow(ViewModel_BaseEventHandler):
    """
    Bảng 3 columns:
    Condition Name | File Name | Categories | Analysis | Preview | Line in Revit | Base Level
    """
    def __init__(self, vm=None):
        ViewModel_BaseEventHandler.__init__(self)
        self._vm              = vm
        self._condition_name  = ''
        self._file_name       = ''
        self._category        = ''
        self._rules           = []       # list[RuleRow] snapshot

        self._analysis_status = 'x'     # 'x' = chưa | 'v' = hoàn thành
        self._preview_checked = False
        self._axis_align_status = ''    # '' = chưa chạy | 'v' = OK | '!' = cần kiểm tra
        self._revit_line_ref  = None    # Revit element (đường tham chiếu cho transform)
        self._base_level      = ''

        self.result_elements  = []      # list[elem] sau Analysis
        self.cad_groups       = []      # list[CadGroup]

        self.color = _next_color()      # màu hiển thị trên Canvas

    # ---- ConditionName ----
    @property
    def ConditionName(self):
        return self._condition_name

    @ConditionName.setter
    def ConditionName(self, v):
        self._condition_name = v
        self.OnPropertyChanged('ConditionName')

    # ---- FileName ----
    @property
    def FileName(self):
        return self._file_name

    @FileName.setter
    def FileName(self, v):
        self._file_name = v
        self.OnPropertyChanged('FileName')

    # ---- Category ----
    @property
    def Category(self):
        return self._category

    @Category.setter
    def Category(self, v):
        self._category = v
        self.OnPropertyChanged('Category')

    # ---- Rules ----
    @property
    def Rules(self):
        return self._rules

    def save_rules_snapshot(self, rule_rows):
        """Lưu snapshot rules vào condition."""
        self._rules = [RuleRow.from_dict(r.to_dict()) for r in rule_rows]

    # ---- AnalysisStatus ----
    @property
    def AnalysisStatus(self):
        return self._analysis_status

    @AnalysisStatus.setter
    def AnalysisStatus(self, v):
        self._analysis_status = v
        self.OnPropertyChanged('AnalysisStatus')
        self.OnPropertyChanged('AnalysisDone')
        self.OnPropertyChanged('CanPreview')

    @property
    def AnalysisDone(self):
        return self._analysis_status == 'v'

    @property
    def CanPreview(self):
        """Checkbox Preview chỉ enable khi Analysis = v."""
        return self.AnalysisDone

    # ---- AxisAlignStatus ----
    @property
    def AxisAlignStatus(self):
        return self._axis_align_status

    @AxisAlignStatus.setter
    def AxisAlignStatus(self, v):
        self._axis_align_status = v or ''
        self.OnPropertyChanged('AxisAlignStatus')
        self.OnPropertyChanged('AxisAlignLabel')
        self.OnPropertyChanged('AxisAlignDone')

    @property
    def AxisAlignDone(self):
        return self._axis_align_status == 'v'

    @property
    def AxisAlignLabel(self):
        if self._axis_align_status == 'v':
            return u'\u2713'                         # ✓
        if self._axis_align_status == '!':
            return u'Vui l\u00f2ng ki\u1ec3m tra AxisAlign'  # warning
        return u''

    # ---- PreviewChecked ----
    @property
    def PreviewChecked(self):
        return self._preview_checked

    @PreviewChecked.setter
    def PreviewChecked(self, v):
        if v and not self.AnalysisDone:
            return  # chặn tick khi chưa analysis
        self._preview_checked = bool(v)
        self.OnPropertyChanged('PreviewChecked')
        if self._vm:
            self._vm.refresh_preview_groups()

    # ---- RevitLineRef ----
    @property
    def RevitLineRef(self):
        return self._revit_line_ref

    @RevitLineRef.setter
    def RevitLineRef(self, v):
        self._revit_line_ref = v
        self.OnPropertyChanged('RevitLineRef')
        self.OnPropertyChanged('RevitLineLabel')

    @property
    def RevitLineLabel(self):
        """Text hiển thị trên button Line in Revit."""
        if self._revit_line_ref is not None:
            return u'\u2713 OK'   # ✓ OK
        return 'Select'

    # ---- BaseLevel ----
    @property
    def BaseLevel(self):
        return self._base_level

    @BaseLevel.setter
    def BaseLevel(self, v):
        self._base_level = v
        self.OnPropertyChanged('BaseLevel')

    # ---- Levels dropdown (từ ViewModel) ----
    @property
    def AvailableLevels(self):
        if self._vm:
            return self._vm.Levels
        return []

    # ---- is_ready – điều kiện để Create Model ----
    def is_ready(self):
        return (
            self._analysis_status == 'v'
            and bool(self._base_level)
            and self._revit_line_ref is not None
            and len(self.result_elements) > 0
            and all(g.is_ready() for g in self.cad_groups)
        )

    # ---- CreateModelStatus ----
    @property
    def CreateModelStatus(self):
        return getattr(self, '_create_model_status', '')

    @CreateModelStatus.setter
    def CreateModelStatus(self, v):
        self._create_model_status = v or ''
        self.OnPropertyChanged('CreateModelStatus')
        self.OnPropertyChanged('CreateModelLabel')
        self.OnPropertyChanged('CreateModelDone')

    @property
    def CreateModelDone(self):
        return getattr(self, '_create_model_status', '') == 'v'

    @property
    def CreateModelLabel(self):
        status = getattr(self, '_create_model_status', '')
        if status == 'v':
            return u'\u2713 Created'
        return u'Create ?'


# =============================================
# ModelByCadViewModel – ViewModel chính
# =============================================
class ModelByCadViewModel(ViewModel_BaseEventHandler):
    """ViewModel chính cho tool Model By CAD."""

    def __init__(self):
        ViewModel_BaseEventHandler.__init__(self)

        # Files đã load: {filename: {'elements': [], 'layers': [], 'path': ''}}
        self.loaded_files = {}

        # Danh sách tên file đã load – persistent ObservableCollection cho binding
        self._loaded_file_names = ObservableCollection[object]()

        # Đường tham chiếu CAD (global, dùng khi tính transform nếu cần)
        self.cad_grid_elements = []

        # Đường tham chiếu đường lưới Revit (global)
        self.revit_grids = []

        # Bảng 3 – danh sách conditions
        self._conditions = ObservableCollection[object]()

        # Bảng 2 – rules đang edit
        self._current_rules = ObservableCollection[object]()

        # Set up data type – groups từ các conditions đang Preview
        self._preview_groups = ObservableCollection[object]()

        # Selection state
        self._selected_condition  = None
        self._selected_group      = None
        self._selected_element    = None
        self._selected_beam_info  = None   # SelectedBeamInfoRow khi click beam trên canvas

        # Revit data
        self._categories        = ObservableCollection[object]()
        self._family_types      = ObservableCollection[object]()
        self._levels            = ObservableCollection[object]()
        self._category_type_map = {}

        # Bảng 1 input fields (binding)
        self._cond_name     = ''
        self._cond_file     = ''
        self._cond_category = ''

        self._status         = ''
        self._cad_grid_layer = ''  # Layer name tham chiếu CAD (thay thế cho interactive select)

    # ──────────────────────────────────────────────────────────
    #   LOADED FILES
    # ──────────────────────────────────────────────────────────
    def add_loaded_file(self, filename, elements, layers, filepath='', texts=None):
        self.loaded_files[filename] = {
            'elements': elements,
            'layers'  : sorted(set(layers)),
            'path'    : filepath,
            'texts'   : texts or [],   # list of (content, x, y, layer)
        }
        # Rebuild persistent collection so ComboBox sees the change
        self._loaded_file_names.Clear()
        for name in sorted(self.loaded_files.keys()):
            self._loaded_file_names.Add(name)
        self.OnPropertyChanged('LoadedFileNames')
        self.OnPropertyChanged('AvailableLayers')

    @property
    def LoadedFileNames(self):
        return self._loaded_file_names

    @property
    def AvailableLayers(self):
        """Layer names của file đang chọn ở Bảng 1."""
        fdata = self.loaded_files.get(self._cond_file, {})
        return fdata.get('layers', [])

    def get_elements_for_file(self, filename):
        return self.loaded_files.get(filename, {}).get('elements', [])

    def get_texts_for_file(self, filename):
        """Trả về list[(content, x, y, layer)] của file đã load."""
        return self.loaded_files.get(filename, {}).get('texts', [])

    # ──────────────────────────────────────────────────────────
    #   CONDITIONS – Bảng 3
    # ──────────────────────────────────────────────────────────
    @property
    def Conditions(self):
        return self._conditions

    @property
    def SelectedCondition(self):
        return self._selected_condition

    @SelectedCondition.setter
    def SelectedCondition(self, v):
        self._selected_condition = v
        self.OnPropertyChanged('SelectedCondition')
        if v is not None:
            # Load lại Bảng 1
            self._cond_name     = v.ConditionName
            self._cond_file     = v.FileName
            self._cond_category = v.Category
            self.OnPropertyChanged('CondName')
            self.OnPropertyChanged('CondFile')
            self.OnPropertyChanged('CondCategory')
            self.OnPropertyChanged('AvailableLayers')
            # Load lại Bảng 2
            self._current_rules.Clear()
            for r in v.Rules:
                self._current_rules.Add(RuleRow.from_dict(r.to_dict()))

    def add_condition(self, name, filename, category, rules_snapshot):
        """Tạo ConditionRow mới và thêm vào Bảng 3."""
        row = ConditionRow(vm=self)
        row.ConditionName = name
        row.FileName      = filename
        row.Category      = category
        row.save_rules_snapshot(rules_snapshot)
        self._conditions.Add(row)
        return row

    def remove_selected_condition(self):
        cond = self._selected_condition
        if cond and cond in list(self._conditions):
            # Xóa groups liên quan khỏi preview
            for g in list(cond.cad_groups):
                if g in list(self._preview_groups):
                    self._preview_groups.Remove(g)
            self._conditions.Remove(cond)
            self._selected_condition = None
            self.OnPropertyChanged('SelectedCondition')

    # ──────────────────────────────────────────────────────────
    #   CURRENT RULES – Bảng 2
    # ──────────────────────────────────────────────────────────
    @property
    def CurrentRules(self):
        return self._current_rules

    def add_rule(self):
        self._current_rules.Add(RuleRow('Layer Name', 'Equal', ''))

    def remove_rule(self, rule):
        if rule and rule in list(self._current_rules):
            self._current_rules.Remove(rule)

    def update_condition_rules(self):
        """
        Ghi CurrentRules vào condition đang chọn → reset analysis.
        Trả về True nếu thành công.
        """
        if self._selected_condition is None:
            return False
        cond = self._selected_condition
        cond.save_rules_snapshot(list(self._current_rules))
        # Rules thay đổi → cần chạy lại analysis
        cond.AnalysisStatus    = 'x'
        cond.PreviewChecked    = False
        cond.AxisAlignStatus   = ''
        cond.CreateModelStatus = ''
        cond.result_elements   = []
        cond.cad_groups        = []
        self.refresh_preview_groups()
        return True

    # ──────────────────────────────────────────────────────────
    #   PREVIEW GROUPS – Set up data type
    # ──────────────────────────────────────────────────────────
    @property
    def PreviewGroups(self):
        return self._preview_groups

    def refresh_preview_groups(self):
        """Gom CadGroup từ tất cả conditions đang PreviewChecked=True."""
        self._preview_groups.Clear()
        for cond in self._conditions:
            if cond.PreviewChecked:
                for g in cond.cad_groups:
                    self._preview_groups.Add(g)

    def get_previewed_elements_with_condition(self):
        """
        Trả về list[(elem, ConditionRow)] cho canvas.
        Dùng condition.color để tô màu theo condition.
        """
        result = []
        for cond in self._conditions:
            if cond.PreviewChecked:
                for elem in cond.result_elements:
                    result.append((elem, cond))
        return result

    # ──────────────────────────────────────────────────────────
    #   SELECTION STATE
    # ──────────────────────────────────────────────────────────
    @property
    def SelectedGroup(self):
        return self._selected_group

    @SelectedGroup.setter
    def SelectedGroup(self, v):
        self._selected_group = v
        self.OnPropertyChanged('SelectedGroup')

    @property
    def SelectedElement(self):
        return self._selected_element

    @SelectedElement.setter
    def SelectedElement(self, v):
        self._selected_element = v
        self.OnPropertyChanged('SelectedElement')

    @property
    def SelectedBeamInfo(self):
        return self._selected_beam_info

    @SelectedBeamInfo.setter
    def SelectedBeamInfo(self, v):
        self._selected_beam_info = v
        self.OnPropertyChanged('SelectedBeamInfo')
        self.OnPropertyChanged('BeamInfoVisible')

    @property
    def BeamInfoVisible(self):
        from System.Windows import Visibility
        return Visibility.Visible if self._selected_beam_info is not None else Visibility.Collapsed

    # ──────────────────────────────────────────────────────────
    #   BẢNG 1 INPUT FIELDS
    # ──────────────────────────────────────────────────────────
    @property
    def CondName(self):
        return self._cond_name

    @CondName.setter
    def CondName(self, v):
        self._cond_name = v
        self.OnPropertyChanged('CondName')

    @property
    def CondFile(self):
        return self._cond_file

    @CondFile.setter
    def CondFile(self, v):
        self._cond_file = v
        self.OnPropertyChanged('CondFile')
        self.OnPropertyChanged('AvailableLayers')

    @property
    def CondCategory(self):
        return self._cond_category

    @CondCategory.setter
    def CondCategory(self, v):
        self._cond_category = v
        self.OnPropertyChanged('CondCategory')

    # ──────────────────────────────────────────────────────────
    #   REVIT DROPDOWNS
    # ──────────────────────────────────────────────────────────
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

    def set_revit_family_types(self, names):
        self._family_types.Clear()
        for n in names:
            self._family_types.Add(n)

    def set_revit_levels(self, names):
        self._levels.Clear()
        for n in names:
            self._levels.Add(n)

    def set_category_type_map(self, mapping):
        self._category_type_map = mapping

    def get_types_for_category(self, category_name):
        return self._category_type_map.get(category_name, [])

    # ──────────────────────────────────────────────────────────
    #   DELETE HELPERS
    # ──────────────────────────────────────────────────────────
    def remove_element(self, elem):
        """Xóa 1 element đơn lẻ (canvas click) – giữ nhóm nếu còn phần tử khác."""
        for cond in self._conditions:
            if elem in cond.result_elements:
                cond.result_elements.remove(elem)
                for g in list(cond.cad_groups):
                    if elem in g.elements:
                        g.elements.remove(elem)
                        g.notify_count_changed()
                        if not g.elements:
                            cond.cad_groups.remove(g)
                            if g in list(self._preview_groups):
                                self._preview_groups.Remove(g)
                        break
                return True
        return False

    def remove_group(self, group):
        """Xóa cả nhóm (1 dòng Set up data type) và tất cả elements của nó."""
        for cond in self._conditions:
            if group in cond.cad_groups:
                for elem in list(group.elements):
                    if elem in cond.result_elements:
                        cond.result_elements.remove(elem)
                group.elements[:] = []
                cond.cad_groups.remove(group)
                if group in list(self._preview_groups):
                    self._preview_groups.Remove(group)
                return True
        return False

    # ──────────────────────────────────────────────────────────
    #   CLEAR ALL
    # ──────────────────────────────────────────────────────────
    def clear_all(self):
        self._conditions.Clear()
        self._current_rules.Clear()
        self._preview_groups.Clear()
        self._selected_condition  = None
        self._selected_group      = None
        self._selected_element    = None
        self._selected_beam_info  = None
        self.cad_grid_elements    = []
        self.revit_grids         = []
        self._cad_grid_layer     = ''
        self.loaded_files.clear()
        self._loaded_file_names.Clear()
        self.OnPropertyChanged('SelectedCondition')
        self.OnPropertyChanged('SelectedGroup')
        self.OnPropertyChanged('SelectedElement')
        self.OnPropertyChanged('SelectedBeamInfo')
        self.OnPropertyChanged('BeamInfoVisible')
        self.OnPropertyChanged('LoadedFileNames')
        self.OnPropertyChanged('AvailableLayers')
        self.OnPropertyChanged('CadGridLayer')

    # ──────────────────────────────────────────────────────────
    #   STATUS
    # ──────────────────────────────────────────────────────────
    @property
    def Status(self):
        return self._status

    @Status.setter
    def Status(self, v):
        self._status = v
        self.OnPropertyChanged('Status')

    # ──────────────────────────────────────────────────────────
    #   CAD GRID LAYER
    # ──────────────────────────────────────────────────────────
    @property
    def CadGridLayer(self):
        return self._cad_grid_layer

    @CadGridLayer.setter
    def CadGridLayer(self, v):
        self._cad_grid_layer = v or ''
        self.OnPropertyChanged('CadGridLayer')
