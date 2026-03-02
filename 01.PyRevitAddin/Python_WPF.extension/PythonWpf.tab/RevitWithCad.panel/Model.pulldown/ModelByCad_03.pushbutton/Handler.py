# -*- coding: utf-8 -*-
"""
Handler.py – Tất cả event handlers và ExternalEvent handlers cho ModelByCad.
"""
import math
import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows.Shapes  import Polyline, Line, Ellipse, Polygon
from System.Windows.Controls import Canvas, TextBlock, Button
from System.Windows.Media   import SolidColorBrush, Colors, Color
from System.Windows         import Point, Thickness
from System.Windows.Forms   import (MessageBox, MessageBoxButtons,
                                    MessageBoxIcon, DialogResult as WFDialogResult,
                                    OpenFileDialog)

from CadUtils import (
    get_acad_doc, load_file_to_doc, extract_all_from_doc,
    select_grid_in_cad,
    filter_elements_by_rules, analyze_condition,
    merge_lines_to_closed_polylines, group_elements_by_label,
    select_beam_elements_in_cad, group_beam_pairs_by_label,
    CadBeamPair, detect_beams_from_lines, BeamAxis,
    _pair_texts_with_beams, align_elements_to_axis,
)
from ViewModel import CadGroup, ConditionRow, RuleRow, SelectedBeamInfoRow

# ─────────────────────────────────────────────────────────
#   REVIT API
# ─────────────────────────────────────────────────────────
try:
    from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
    from Autodesk.Revit.UI.Selection import ObjectType
    from Autodesk.Revit.DB import (
        Grid, FilteredElementCollector, Level, FamilySymbol,
        Transaction, XYZ, Line as RvtLine, BuiltInParameter
    )
    from Autodesk.Revit.DB.Structure import StructuralType
    _REVIT_AVAILABLE = True
except Exception:
    _REVIT_AVAILABLE = False


# ─────────────────────────────────────────────────────────
#   GEOMETRY TRANSFORM HELPERS  (CAD mm → Revit feet)
# ─────────────────────────────────────────────────────────
_MM_TO_FEET = 1.0 / 304.8


def _get_cad_ref(cad_elem):
    s = cad_elem.points[0]
    e = cad_elem.points[-1]
    dx, dy = e[0]-s[0], e[1]-s[1]
    L = math.sqrt(dx*dx + dy*dy) or 1.0
    return (s[0], s[1], dx/L, dy/L)


def _get_revit_ref(revit_elem):
    try:
        curve = revit_elem.Curve
    except Exception:
        curve = revit_elem.Location.Curve
    start = curve.GetEndPoint(0)
    d     = curve.Direction
    return (start.X, start.Y, d.X, d.Y)


def _build_transform(cad_elem, revit_elem):
    sx_c, sy_c, dx_c, dy_c = _get_cad_ref(cad_elem)
    sx_r, sy_r, dx_r, dy_r = _get_revit_ref(revit_elem)
    theta = math.atan2(dy_r, dx_r) - math.atan2(dy_c, dx_c)
    return (sx_c, sy_c), (sx_r, sy_r), theta


def _rotate2d(dx, dy, theta):
    c, s = math.cos(theta), math.sin(theta)
    return dx*c - dy*s, dx*s + dy*c


def _transform_to_revit(cx_mm, cy_mm, s_cad, s_revit_ft, theta):
    dx_mm = cx_mm - s_cad[0]
    dy_mm = cy_mm - s_cad[1]
    dx_r, dy_r = _rotate2d(dx_mm, dy_mm, theta)
    return XYZ(s_revit_ft[0] + dx_r*_MM_TO_FEET,
               s_revit_ft[1] + dy_r*_MM_TO_FEET, 0.0)


def _get_element_center(elem):
    if elem.type == 'circle' and elem.center:
        return elem.center
    if elem.type in ('polyline', 'line', 'beam_line', 'beam_axis') and getattr(elem, 'points', []):
        pts = elem.points
        xs  = [p[0] for p in pts]
        ys  = [p[1] for p in pts]
        return ((min(xs)+max(xs))/2.0, (min(ys)+max(ys))/2.0)
    if elem.type == 'arc' and elem.center:
        return elem.center
    return None


_CAT_TO_STRUCTURAL = {}


def _init_structural_types():
    if not _REVIT_AVAILABLE:
        return
    _CAT_TO_STRUCTURAL['structural columns']      = StructuralType.Column
    _CAT_TO_STRUCTURAL['structural foundations']  = StructuralType.Footing
    _CAT_TO_STRUCTURAL['structural framing']      = StructuralType.Beam
    _CAT_TO_STRUCTURAL['structural framings']     = StructuralType.Beam


if _REVIT_AVAILABLE:
    _init_structural_types()


def _category_to_structural_type(cat_name):
    return _CAT_TO_STRUCTURAL.get(cat_name.lower().strip(), StructuralType.NonStructural)


def _resolve_symbols(doc):
    result = {}
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            fam_name = ''
            try: fam_name = sym.FamilyName or ''
            except Exception: pass
            if not fam_name:
                try:
                    if sym.Family: fam_name = sym.Family.Name or ''
                except Exception: pass
            if not fam_name: continue

            sym_name = ''
            try: sym_name = sym.Name or ''
            except Exception: pass
            if not sym_name:
                try:
                    p = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if p: sym_name = p.AsString() or ''
                except Exception: pass
            if not sym_name: continue

            result['{} : {}'.format(fam_name, sym_name)] = sym
        except Exception:
            continue
    return result


def _resolve_levels(doc):
    result = {}
    for lv in FilteredElementCollector(doc).OfClass(Level):
        try: result[lv.Name] = lv
        except Exception: pass
    return result


# ─────────────────────────────────────────────────────────
#   REVIT EXTERNAL EVENT: Select Line for Condition
# ─────────────────────────────────────────────────────────
class SelectRevitLineHandler(IExternalEventHandler if _REVIT_AVAILABLE else object):
    """
    Cho người dùng pick 1 đường trong Revit, gán vào condition đang chờ.
    Window gọi self.pending_condition = <ConditionRow> rồi raise event.
    """
    def __init__(self, vm):
        self.vm                = vm
        self.pending_condition = None   # ConditionRow đang chờ

    def Execute(self, uiapp):
        if self.pending_condition is None:
            return
        cond = self.pending_condition
        try:
            uidoc = uiapp.ActiveUIDocument
            ref   = uidoc.Selection.PickObject(
                ObjectType.Element,
                u"Chon duong tham chieu Revit cho: {}".format(cond.ConditionName)
            )
            elem = uidoc.Document.GetElement(ref.ElementId)
            # Kiểm tra có Curve không
            try:
                _get_revit_ref(elem)
            except Exception:
                print("SelectRevitLineHandler: element khong co Curve.")
                return
            cond.RevitLineRef = elem
            self.vm.Status = u"[{}] Da chon duong tham chieu Revit.".format(cond.ConditionName)
        except Exception as ex:
            print("SelectRevitLineHandler error: {}".format(ex))
        finally:
            self.pending_condition = None

    def GetName(self):
        return "SelectRevitLine"


# ─────────────────────────────────────────────────────────
#   REVIT EXTERNAL EVENT: Select Global CAD Grid Reference
# ─────────────────────────────────────────────────────────
class SelectRevitGridHandler(IExternalEventHandler if _REVIT_AVAILABLE else object):
    """Chọn đường tham chiếu Revit toàn cục (dùng khi cần 1 reference chung)."""
    def __init__(self, vm):
        self.vm = vm

    def Execute(self, uiapp):
        try:
            uidoc = uiapp.ActiveUIDocument
            ref   = uidoc.Selection.PickObject(
                ObjectType.Element,
                u"Chon duong tham chieu Revit (global)"
            )
            elem = uidoc.Document.GetElement(ref.ElementId)
            try: _get_revit_ref(elem)
            except Exception:
                print("SelectRevitGridHandler: khong co Curve.")
                return
            self.vm.revit_grids = [elem]
            self.vm.Status = u"Da chon duong tham chieu Revit (global)."
        except Exception as ex:
            print("SelectRevitGridHandler error: {}".format(ex))

    def GetName(self):
        return "SelectRevitGrid"


# ─────────────────────────────────────────────────────────
#   REVIT EXTERNAL EVENT: Create Model
# ─────────────────────────────────────────────────────────
class CreateModelHandler(IExternalEventHandler if _REVIT_AVAILABLE else object):
    """
    Tạo FamilyInstance trong Revit cho tất cả conditions đã đủ điều kiện.
    Mỗi condition dùng cặp (CAD ref global, Revit line của condition đó).
    """
    def __init__(self, vm):
        self.vm = vm

    def Execute(self, uiapp):
        vm  = self.vm
        doc = uiapp.ActiveUIDocument.Document

        # Lấy layer name tham chiếu CAD (từ TextBox TxtGridLayer)
        grid_layer = (getattr(vm, 'CadGridLayer', '') or '').strip().upper()
        if not grid_layer:
            MessageBox.Show(
                u"Chua nhap ten layer tham chieu CAD (Grid Layer).\n"
                u"Vui long nhap layer name vao o 'Grid Layer:' tren thanh cong cu.",
                u"Thieu tham chieu CAD"
            )
            return

        # Lọc conditions đủ điều kiện
        ready = [c for c in vm.Conditions if c.is_ready()]
        if not ready:
            missing = []
            for c in vm.Conditions:
                if c.AnalysisStatus != 'v':
                    missing.append(u"  [{}] Chua Analysis".format(c.ConditionName))
                elif not c.cad_groups:
                    missing.append(u"  [{}] Khong co ket qua".format(c.ConditionName))
                elif not c.BaseLevel:
                    missing.append(u"  [{}] Chua chon Base Level".format(c.ConditionName))
                elif c.RevitLineRef is None:
                    missing.append(u"  [{}] Chua chon Line in Revit".format(c.ConditionName))
                elif not all(g.is_ready() for g in c.cad_groups):
                    missing.append(u"  [{}] Chua map du Family Type".format(c.ConditionName))
            MessageBox.Show(
                u"Khong co condition nao du dieu kien:\n" + u"\n".join(missing[:8]),
                u"Thieu thong tin"
            )
            return

        symbol_map = _resolve_symbols(doc)
        level_map  = _resolve_levels(doc)

        total_created = 0
        total_skipped = 0
        all_errors    = []

        t = Transaction(doc, u"Create Model from CAD")
        t.Start()
        try:
            for cond in ready:
                # Tìm element tham chiếu từ file của condition theo layer
                file_elements = vm.get_elements_for_file(cond.FileName)
                grid_elem = None
                for fe in file_elements:
                    if (getattr(fe, 'layer', '').upper() == grid_layer and
                            fe.type in ('line', 'polyline') and
                            len(getattr(fe, 'points', [])) >= 2):
                        grid_elem = fe
                        break
                if grid_elem is None:
                    all_errors.append(
                        u"[{}] Khong tim thay duong tham chieu tren layer '{}' trong file '{}'".format(
                            cond.ConditionName, grid_layer, cond.FileName))
                    continue

                # Build transform: CAD layer ref → Revit condition ref
                try:
                    s_cad, s_revit_ft, theta = _build_transform(grid_elem, cond.RevitLineRef)
                except Exception as ex:
                    all_errors.append(u"[{}] Loi transform: {}".format(cond.ConditionName, ex))
                    continue

                level = level_map.get(cond.BaseLevel)
                if level is None:
                    all_errors.append(u"[{}] Khong tim thay Level '{}'".format(
                        cond.ConditionName, cond.BaseLevel))
                    continue

                for group in cond.cad_groups:
                    if not group.is_ready():
                        total_skipped += len(group.elements)
                        continue

                    sym_key = group.FamilyType
                    symbol  = symbol_map.get(sym_key)
                    if symbol is None:
                        total_skipped += len(group.elements)
                        all_errors.append(u"[{}] Khong tim thay FamilySymbol: '{}'".format(
                            cond.ConditionName, sym_key))
                        continue

                    if not symbol.IsActive:
                        symbol.Activate()
                        doc.Regenerate()

                    struct_type = _category_to_structural_type(cond.Category)
                    level_z     = level.Elevation

                    for elem in group.elements:
                        try:
                            # Per-element FamilyType override
                            fam_override = getattr(elem, 'family_type_override', None)
                            elem_symbol  = symbol_map.get(fam_override, symbol) if fam_override else symbol
                            if elem_symbol is None:
                                elem_symbol = symbol

                            if isinstance(elem, (CadBeamPair, BeamAxis)):
                                # Dầm: curve overload (start/end đã là tuple (x,y))
                                xyz_s = _transform_to_revit(
                                    elem.start[0], elem.start[1], s_cad, s_revit_ft, theta)
                                xyz_e = _transform_to_revit(
                                    elem.end[0],   elem.end[1],   s_cad, s_revit_ft, theta)
                                xyz_s = XYZ(xyz_s.X, xyz_s.Y, level_z)
                                xyz_e = XYZ(xyz_e.X, xyz_e.Y, level_z)
                                curve = RvtLine.CreateBound(xyz_s, xyz_e)
                                inst  = doc.Create.NewFamilyInstance(
                                    curve, elem_symbol, level, StructuralType.Beam)
                                try:
                                    # Per-element override có ưu tiên; fallback: edge→Left(0), other→Center(2)
                                    loc_override = getattr(elem, 'location_type_override', None)
                                    if loc_override is not None:
                                        y_just = loc_override
                                    else:
                                        is_edge = (isinstance(elem, BeamAxis) and
                                                   getattr(elem, 'beam_type', '') == 'edge')
                                        y_just = 0 if is_edge else 2
                                    p = inst.get_Parameter(BuiltInParameter.Y_JUSTIFICATION)
                                    if p and not p.IsReadOnly: p.Set(y_just)
                                except Exception: pass
                            else:
                                # Cột / Móng: point overload
                                center = _get_element_center(elem)
                                if center is None:
                                    total_skipped += 1
                                    continue
                                xyz = _transform_to_revit(
                                    center[0], center[1], s_cad, s_revit_ft, theta)
                                xyz = XYZ(xyz.X, xyz.Y, level_z)
                                doc.Create.NewFamilyInstance(
                                    xyz, elem_symbol, level, struct_type)
                            total_created += 1
                        except Exception as ex_elem:
                            total_skipped += 1
                            all_errors.append(str(ex_elem))

            t.Commit()
        except Exception as ex_t:
            t.RollBack()
            print(u"CreateModel: Transaction error, rolled back: {}".format(ex_t))
            MessageBox.Show(u"Loi transaction: {}".format(ex_t), u"Loi")
            return

        vm.Status = u"Tao xong: {} instances. Bo qua: {}.".format(total_created, total_skipped)
        print(u"CreateModel: {} created, {} skipped.".format(total_created, total_skipped))
        for msg in all_errors:
            print(u"  [WARN] {}".format(msg))

        report = u"Da tao thanh cong {} instances!".format(total_created)
        if total_skipped:
            report += u"\nBo qua: {} elements.".format(total_skipped)
        if all_errors:
            report += u"\n\nChi tiet:\n" + u"\n".join(all_errors[:5])
        MessageBox.Show(report, u"Ket qua Create Model")

    def GetName(self):
        return "CreateModel"


# ─────────────────────────────────────────────────────────
#   REVIT EXTERNAL EVENT: Create Model for Single Condition
# ─────────────────────────────────────────────────────────
class CreateModelSingleHandler(IExternalEventHandler if _REVIT_AVAILABLE else object):
    """
    Tạo FamilyInstance trong Revit chỉ cho 1 ConditionRow được chỉ định.
    window đặt self.pending_condition rồi raise event.
    """
    def __init__(self, vm):
        self.vm                = vm
        self.pending_condition = None   # ConditionRow cần create
        self._window           = None   # set bởi script.py sau khi khởi tạo

    def Execute(self, uiapp):
        cond = self.pending_condition
        if cond is None:
            return
        self.pending_condition = None

        vm  = self.vm
        doc = uiapp.ActiveUIDocument.Document

        # Lấy layer name tham chiếu CAD
        grid_layer = (getattr(vm, 'CadGridLayer', '') or '').strip().upper()
        if not grid_layer:
            MessageBox.Show(
                u"Chưa nhập tên layer tham chiếu CAD (Grid Layer).",
                u"Thiếu tham chiếu CAD"
            )
            return

        # Tìm element tham chiếu trên layer grid
        file_elements = vm.get_elements_for_file(cond.FileName)
        grid_elem = None
        for fe in file_elements:
            if (getattr(fe, 'layer', '').upper() == grid_layer and
                    fe.type in ('line', 'polyline') and
                    len(getattr(fe, 'points', [])) >= 2):
                grid_elem = fe
                break
        if grid_elem is None:
            MessageBox.Show(
                u"Không tìm thấy đường tham chiếu trên layer '{}' trong file '{}'.".format(
                    grid_layer, cond.FileName),
                u"Thiếu tham chiếu CAD"
            )
            return

        symbol_map = _resolve_symbols(doc)
        level_map  = _resolve_levels(doc)

        try:
            s_cad, s_revit_ft, theta = _build_transform(grid_elem, cond.RevitLineRef)
        except Exception as ex:
            MessageBox.Show(u"Lỗi transform: {}".format(ex), u"Create Model")
            return

        level = level_map.get(cond.BaseLevel)
        if level is None:
            MessageBox.Show(
                u"Không tìm thấy Level '{}'.".format(cond.BaseLevel),
                u"Create Model"
            )
            return

        struct_type = _category_to_structural_type(cond.Category)
        level_z     = level.Elevation
        created     = 0
        skipped     = 0
        errors      = []

        t = Transaction(doc, u"Create Model [{}]".format(cond.ConditionName))
        t.Start()
        try:
            for group in cond.cad_groups:
                if not group.is_ready():
                    skipped += len(group.elements)
                    continue
                symbol = symbol_map.get(group.FamilyType)
                if symbol is None:
                    skipped += len(group.elements)
                    errors.append(u"Không tìm thấy FamilySymbol: '{}'".format(group.FamilyType))
                    continue
                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()

                for elem in group.elements:
                    try:
                        # Per-element FamilyType override
                        fam_override = getattr(elem, 'family_type_override', None)
                        elem_symbol  = symbol_map.get(fam_override, symbol) if fam_override else symbol
                        if elem_symbol is None:
                            elem_symbol = symbol

                        if isinstance(elem, (CadBeamPair, BeamAxis)):
                            xyz_s = _transform_to_revit(elem.start[0], elem.start[1], s_cad, s_revit_ft, theta)
                            xyz_e = _transform_to_revit(elem.end[0],   elem.end[1],   s_cad, s_revit_ft, theta)
                            xyz_s = XYZ(xyz_s.X, xyz_s.Y, level_z)
                            xyz_e = XYZ(xyz_e.X, xyz_e.Y, level_z)
                            curve = RvtLine.CreateBound(xyz_s, xyz_e)
                            inst  = doc.Create.NewFamilyInstance(
                                curve, elem_symbol, level, StructuralType.Beam)
                            try:
                                # Per-element override có ưu tiên; fallback: edge→Left(0), other→Center(2)
                                loc_override = getattr(elem, 'location_type_override', None)
                                if loc_override is not None:
                                    y_just = loc_override
                                else:
                                    is_edge = (isinstance(elem, BeamAxis) and
                                               getattr(elem, 'beam_type', '') == 'edge')
                                    y_just = 0 if is_edge else 2
                                p = inst.get_Parameter(BuiltInParameter.Y_JUSTIFICATION)
                                if p and not p.IsReadOnly: p.Set(y_just)
                            except Exception: pass
                        else:
                            center = _get_element_center(elem)
                            if center is None:
                                skipped += 1
                                continue
                            xyz = _transform_to_revit(center[0], center[1], s_cad, s_revit_ft, theta)
                            xyz = XYZ(xyz.X, xyz.Y, level_z)
                            doc.Create.NewFamilyInstance(xyz, elem_symbol, level, struct_type)
                        created += 1
                    except Exception as ex_e:
                        skipped += 1
                        errors.append(str(ex_e))
            t.Commit()
        except Exception as ex_t:
            t.RollBack()
            MessageBox.Show(u"Lỗi transaction: {}".format(ex_t), u"Create Model")
            return

        cond.CreateModelStatus = 'v'
        vm.Status = u"[{}] Tạo xong: {} instances. Bỏ qua: {}.".format(
            cond.ConditionName, created, skipped)

        report = u"[{}] Tạo thành công {} instances!".format(cond.ConditionName, created)
        if skipped:
            report += u"\nBỏ qua: {} elements.".format(skipped)
        if errors:
            report += u"\n\nChi tiết:\n" + u"\n".join(errors[:5])
        MessageBox.Show(report, u"Kết quả Create Model")

    def GetName(self):
        return "CreateModelSingle"


# ─────────────────────────────────────────────────────────
#   BIND ALL HANDLERS
# ─────────────────────────────────────────────────────────
def bind_handlers(window):
    """Gắn tất cả event handlers vào window."""
    # Top bar buttons
    window.BtnLoadFile.Tag        = window
    window.BtnLoadFile.Click     += on_load_file

    window.BtnDelete.Tag   = window
    window.BtnDelete.Click += on_delete

    window.BtnRefresh.Tag   = window
    window.BtnRefresh.Click += on_refresh

    # Bảng 1 + 2 buttons
    window.BtnRemoveCondition.Tag   = window
    window.BtnRemoveCondition.Click += on_remove_condition

    window.BtnCopyCondition.Tag   = window
    window.BtnCopyCondition.Click += on_copy_condition

    window.BtnAddRule.Tag   = window
    window.BtnAddRule.Click += on_add_rule

    window.BtnRemoveRule.Tag   = window
    window.BtnRemoveRule.Click += on_remove_rule

    window.BtnUpdateCond.Tag   = window
    window.BtnUpdateCond.Click += on_update_condition

    # Bảng 3 DataGrid – intercept cell-template button clicks
    window.DgConditions.Tag = window
    window.DgConditions.PreviewMouseLeftButtonDown += on_conditions_datagrid_click
    # Redraw canvas when any cell is clicked (catches checkbox Preview toggle)
    window.DgConditions.MouseLeftButtonUp += on_conditions_lbup

    # Bảng 3 — load Bảng 1+2 khi chọn row
    window.DgConditions.SelectionChanged += on_conditions_selection_changed

    # Set up data type DataGrid
    window.DgGroups.Tag = window
    window.DgGroups.SelectionChanged += on_groups_selection_changed

    # Canvas
    window.DrawingCanvas.Tag = window
    window.DrawingCanvas.MouseLeftButtonDown += on_canvas_click
    window.DrawingCanvas.SizeChanged         += on_canvas_size_changed


# ─────────────────────────────────────────────────────────
#   TOP BAR HANDLERS
# ─────────────────────────────────────────────────────────
def on_load_file(sender, e):
    """
    Load 1 hoac nhieu file DWG vao ViewModel.
    Voi moi file: ket noi AutoCAD -> mo file -> extract layers + elements -> cache.
    """
    try:
        window = sender.Tag
        vm     = window.DataContext

        dlg = OpenFileDialog()
        dlg.Title       = u"Chon file DWG"
        dlg.Filter      = u"AutoCAD files (*.dwg)|*.dwg|All files (*.*)|*.*"
        dlg.Multiselect = True
        if dlg.ShowDialog() != WFDialogResult.OK:
            return

        files  = list(dlg.FileNames)
        loaded = []
        failed = []
        for filepath in files:
            try:
                doc = load_file_to_doc(filepath)
                if doc is None:
                    failed.append(os.path.basename(filepath))
                    continue
                elements, layers, texts = extract_all_from_doc(doc)
                filename = os.path.basename(filepath)
                vm.add_loaded_file(filename, elements, layers, filepath, texts)
                loaded.append(u"{} ({} elems, {} layers)".format(
                    filename, len(elements), len(layers)))
            except Exception as ex:
                failed.append(os.path.basename(filepath))
                print(u"on_load_file error [{}]: {}".format(filepath, ex))

        if loaded:
            msg = u"Da load: " + u", ".join(loaded)
            if failed:
                msg += u"  |  Loi: " + u", ".join(failed)
            vm.Status = msg
        else:
            MessageBox.Show(u"Khong the load bat ky file nao.", u"Load File CAD")

    except Exception as ex:
        MessageBox.Show(u"on_load_file loi: {}".format(ex), u"Load File CAD")


def on_select_grid_cad(sender, e):
    """Chọn đường tham chiếu CAD toàn cục (interactive từ AutoCAD)."""
    window = sender.Tag
    vm     = window.DataContext

    doc = get_acad_doc()
    if not doc:
        MessageBox.Show(u"Khong the ket noi AutoCAD. Hay mo AutoCAD truoc.", u"Loi")
        return

    grids = select_grid_in_cad(doc)
    if not grids:
        MessageBox.Show(u"Can chon dung 1 duong tham chieu trong CAD.", u"Chon tham chieu")
        return

    vm.cad_grid_elements = grids[:1]
    vm.Status = u"Da chon tham chieu CAD. StartPoint: ({:.1f}, {:.1f}) mm".format(
        grids[0].points[0][0], grids[0].points[0][1]
    )
    _redraw(window)


def on_select_grid_revit(sender, e):
    """Raise ExternalEvent chọn đường tham chiếu Revit toàn cục."""
    window = sender.Tag
    if hasattr(window, '_ext_select_grid'):
        window._ext_select_grid.Raise()
    else:
        MessageBox.Show(u"ExternalEvent chua khoi tao.", u"Loi")


def on_delete(sender, e):
    """
    Xóa element đang chọn (canvas click) hoặc group đang chọn (Set up data type).
    Logic giống code cũ.
    """
    window    = sender.Tag
    vm        = window.DataContext
    sel_elem  = getattr(vm, 'SelectedElement', None)
    sel_group = getattr(vm, 'SelectedGroup', None)

    if sel_elem is not None:
        vm.remove_element(sel_elem)
        vm.SelectedElement = None
        vm.SelectedGroup   = None
        vm.Status = u"Da xoa 1 element."
    elif sel_group is not None:
        n = len(sel_group.elements)
        vm.remove_group(sel_group)
        vm.SelectedElement = None
        vm.SelectedGroup   = None
        vm.Status = u"Da xoa nhom '{}' ({} elements).".format(sel_group.label, n)
    else:
        MessageBox.Show(
            u"Chua chon element hay nhom.\n"
            u"- Click shape trong Preview → chon element don le.\n"
            u"- Click dong trong 'Set up data type' → chon ca nhom.",
            u"Delete Selected"
        )
        return

    _redraw(window)


def on_create_model(sender, e):
    """Raise ExternalEvent tạo model trong Revit (top button – legacy)."""
    window = sender.Tag
    if hasattr(window, '_ext_create_model'):
        window._ext_create_model.Raise()
    else:
        MessageBox.Show(u"ExternalEvent chua khoi tao.", u"Loi")


def on_create_model_for_condition(cond, window):
    """
    Tạo model trong Revit chỉ cho 1 condition cụ thể.
    Kiểm tra điều kiện, sau đó raise ExternalEvent với condition đó.
    """
    vm = window.DataContext

    if not cond.is_ready():
        missing = []
        if cond.AnalysisStatus != 'v':
            missing.append(u'- Chưa chạy Analysis')
        if not cond.cad_groups:
            missing.append(u'- Không có kết quả phân tích')
        if not cond.BaseLevel:
            missing.append(u'- Chưa chọn Base Level')
        if cond.RevitLineRef is None:
            missing.append(u'- Chưa chọn Line in Revit')
        if not all(g.is_ready() for g in cond.cad_groups):
            missing.append(u'- Chưa map đủ Family Type trong Set up data type')
        MessageBox.Show(
            u'Condition [{}] chưa đủ điều kiện:\n{}'.format(
                cond.ConditionName, u'\n'.join(missing)),
            u'Create Model'
        )
        return

    if hasattr(window, '_ext_create_model_single'):
        window._ext_create_model_single.pending_condition = cond
        window._ext_create_model_single_event.Raise()
    else:
        # Fallback: call directly (ngoài Revit context, dùng cho debug)
        MessageBox.Show(u'ExternalEvent SingleCreate chưa khởi tạo.', u'Lỗi')


def on_refresh(sender, e):
    """Xóa toàn bộ dữ liệu."""
    res = MessageBox.Show(
        u"Xoa toan bo du lieu (files, conditions, preview)?",
        u"Refresh",
        MessageBoxButtons.YesNo, MessageBoxIcon.Question
    )
    if res != WFDialogResult.Yes:
        return
    window = sender.Tag
    vm     = window.DataContext
    vm.clear_all()
    vm.Status = u"Da xoa toan bo du lieu."
    _redraw(window)


# ─────────────────────────────────────────────────────────
#   BẢNG 1 BUTTONS
# ─────────────────────────────────────────────────────────
def on_insert_condition(sender, e):
    """
    Insert: lấy Bảng 1 + Bảng 2 hiện tại → thêm 1 dòng vào Bảng 3.
    Trình tự:
      1. Lưu rules hiện tại vào condition mới.
      2. Auto-select dòng vừa thêm (setter load rules về Bảng 2).
      3. Populate default template theo category lên Bảng 2
         → sẵn sàng cho lần Insert tiếp theo với category bất kỳ.
    """
    window = sender.Tag
    vm     = window.DataContext

    name     = vm.CondName.strip()
    filename = vm.CondFile
    category = vm.CondCategory

    if not name:
        MessageBox.Show(u"Vui long nhap Condition Name.", u"Thieu thong tin")
        return
    if not filename:
        MessageBox.Show(u"Vui long chon File Name.", u"Thieu thong tin")
        return
    if not category:
        MessageBox.Show(u"Vui long chon Category.", u"Thieu thong tin")
        return

    # Bước 1: Lưu rules user đang nhập vào condition mới
    cond = vm.add_condition(name, filename, category, list(vm.CurrentRules))
    cond.save_rules_snapshot(list(vm.CurrentRules))

    # Bước 2: Auto-select dòng vừa thêm trong Bảng 3
    # (setter tự load rules của cond về Bảng 2 – hiện đúng những gì vừa lưu)
    vm.SelectedCondition = cond
    window.DgConditions.ScrollIntoView(cond)

    vm.Status = u"Da them condition '{}' [{} | {}].".format(name, filename, category)


def on_category_changed(sender, e):
    """
    Khi user thay đổi Category ở Bảng 1,
    populate ngay default rules vào Bảng 2 theo category mới.
    """
    try:
        window = sender.Tag
        vm     = window.DataContext
        category = vm.CondCategory
        if category:
            _populate_default_rules(vm, category)
    except Exception as ex:
        print("on_category_changed error: {}".format(ex))


def _populate_default_rules(vm, category):
    """Cap nhat Bang 2 voi rules mac dinh theo category."""
    vm.CurrentRules.Clear()
    for r in vm._build_default_rules(category):
        vm.CurrentRules.Add(r)


def on_remove_condition(sender, e):
    """Xoa condition dang chon o Bang 3."""
    window = sender.Tag
    vm     = window.DataContext
    if vm.SelectedCondition is None:
        MessageBox.Show(u"Chua chon condition nao trong Bang 3.", u"Remove")
        return
    vm.remove_selected_condition()
    vm.Status = u"Da xoa condition."
    _redraw(window)


def on_copy_condition(sender, e):
    """Tao ban sao cua condition dang chon o Bang 3 (bao gom rules)."""
    window = sender.Tag
    vm     = window.DataContext
    if vm.SelectedCondition is None:
        MessageBox.Show(u"Chua chon condition nao trong Bang 3.", u"Copy")
        return
    source_name = vm.SelectedCondition.ConditionName
    new_row     = vm.copy_condition(vm.SelectedCondition)
    vm.SelectedCondition = new_row
    window.DgConditions.ScrollIntoView(new_row)
    vm.Status = u"Da copy '{}' -> '{}'.".format(source_name, new_row.ConditionName)


# ─────────────────────────────────────────────────────────
#   BẢNG 2 BUTTONS
# ─────────────────────────────────────────────────────────
def on_add_rule(sender, e):
    """Thêm 1 RuleRow mặc định vào Bảng 2."""
    window = sender.Tag
    window.DataContext.add_rule()


def on_remove_rule(sender, e):
    """Xóa RuleRow đang chọn trong Bảng 2."""
    window = sender.Tag
    vm     = window.DataContext
    selected = window.DgRules.SelectedItem
    if selected is None:
        MessageBox.Show(u"Chua chon rule nao.", u"Del Rule")
        return
    vm.remove_rule(selected)


def on_update_condition(sender, e):
    """Ghi rules hiện tại (Bảng 2) vào condition đang chọn (Bảng 3)."""
    window = sender.Tag
    vm     = window.DataContext
    if not vm.update_condition_rules():
        MessageBox.Show(
            u"Chua chon condition nao trong Bang 3.\n"
            u"Click vao dong condition truoc khi Update.",
            u"Update"
        )
        return
    vm.Status = u"Da cap nhat rules cho '{}'.".format(
        vm.SelectedCondition.ConditionName if vm.SelectedCondition else '?')


# ─────────────────────────────────────────────────────────
#   BẢNG 3 – DATAGRID CLICK HANDLER
# ─────────────────────────────────────────────────────────
def on_conditions_datagrid_click(sender, e):
    """
    Intercept click trên các button trong DataGrid Bảng 3:
      - Analysis button  → chạy analysis cho condition đó
      - Line in Revit button → raise ExternalEvent chọn line
      - Create Model button → tạo model cho condition đó
    """
    try:
        source = e.OriginalSource
        # Đi lên cây visual để tìm Button
        btn  = None
        elem = source
        for _ in range(8):
            if isinstance(elem, Button):
                btn = elem
                break
            parent = getattr(elem, 'Parent', None) or getattr(elem, 'TemplatedParent', None)
            if parent is None:
                break
            elem = parent

        if btn is None:
            return

        tag = getattr(btn, 'Tag', None)
        if not isinstance(tag, ConditionRow):
            return

        content = str(btn.Content or '').strip()
        window  = sender.Tag

        # Analysis button (content = 'x' or 'v')
        if content in ('x', 'v'):
            e.Handled = True
            on_run_analysis(tag, window)
            return

        # Line in Revit button (content = 'Select' or '✓ OK')
        if content in ('Select', u'\u2713 OK'):
            e.Handled = True
            on_select_line_for_condition(tag, window)
            return

        # Create Model button (content = 'Create ?' or '✓ Created')
        if content in (u'Create ?', u'\u2713 Created'):
            e.Handled = True
            on_create_model_for_condition(tag, window)
            return

    except Exception as ex:
        print("on_conditions_datagrid_click error: {}".format(ex))


def on_conditions_selection_changed(sender, e):
    """Click dòng Bảng 3 → load lại Bảng 1 + Bảng 2."""
    try:
        window = getattr(sender, 'Tag', None)
        if window is None: return
        vm  = window.DataContext
        sel = sender.SelectedItem
        if sel is None:
            # Deselect → clear SelectedCondition để on_category_changed
            # biết đang ở trạng thái nhập mới và được populate rules
            vm.SelectedCondition = None
            return
        vm.SelectedCondition = sel    # setter tự load Bảng 1 + Bảng 2
    except Exception as ex:
        print("on_conditions_selection_changed error: {}".format(ex))


def on_conditions_lbup(sender, e):
    """MouseLeftButtonUp trên DgConditions: sau khi WPF cập nhật checkbox Preview
    thì refresh preview groups và vẽ lại canvas."""
    try:
        window = getattr(sender, 'Tag', None)
        if window is None: return
        vm = window.DataContext
        vm.refresh_preview_groups()
        _redraw(window)
    except Exception as ex:
        print("on_conditions_lbup error: {}".format(ex))


# ─────────────────────────────────────────────────────────
#   ANALYSIS – chạy phân tích cho 1 condition
# ─────────────────────────────────────────────────────────
def on_run_analysis(cond, window):
    """
    Chạy Analysis cho 1 ConditionRow:
      1. Lấy elements từ file đã load
      2. Filter theo rules (OR logic)
      3. Analyze theo Category
      4. Merge lines → closed poly (nếu là Columns / Walls)
      5. Group → CadGroups
      6. Cập nhật status 'v' hoặc giữ 'x'
    """
    vm = window.DataContext

    filename = cond.FileName
    category = cond.Category or ''
    rules    = cond.Rules

    all_elements = vm.get_elements_for_file(filename)
    if not all_elements:
        MessageBox.Show(
            u"File '{}' chua duoc load hoac khong co elements.\n"
            u"Nhan 'Load File CAD' truoc.".format(filename),
            u"Analysis"
        )
        return

    # 1) Trích tham số của thuật toán từ rules (không phải filter element)
    min_length = 1000.0   # chiều dài tối thiểu sau merge (default)
    min_d      = 100.0    # khoảng cách dầm nhỏ nhất
    max_d      = 1000.0   # khoảng cách dầm lớn nhất
    for r in (rules or []):
        rd = r.to_dict() if hasattr(r, 'to_dict') else r
        param = rd.get('parameter', '')
        try:
            v = float(rd.get('value', 0))
        except (ValueError, TypeError):
            v = 0.0
        if param == 'Length' and rd.get('ruler') == 'is greater than':
            min_length = v
        elif param == 'Min Beam Distance':
            min_d = v
        elif param == 'Max Beam Distance':
            max_d = v

    # 2) Filter theo rules (Layer Name + Length – bỏ qua Min/Max Beam Distance)
    if rules:
        filtered = filter_elements_by_rules(all_elements, rules)
    else:
        filtered = list(all_elements)

    if not filtered:
        cond.AnalysisStatus = 'x'
        MessageBox.Show(
            u"Khong tim thay elements thoa rules.\n"
            u"Kiem tra lai cac rules trong Bang 2.",
            u"Analysis – Khong co ket qua"
        )
        return

    # 3) Analyze theo Category
    cat_lower = category.lower()
    if 'column' in cat_lower or 'foundation' in cat_lower or 'footing' in cat_lower:
        merged   = merge_lines_to_closed_polylines(filtered)
        analyzed = analyze_condition(merged, category)
    elif 'framing' in cat_lower or 'beam' in cat_lower:
        # Truyền tham số người dùng vào thuật toán phát hiện dầm
        analyzed = detect_beams_from_lines(
            filtered,
            min_d           = min_d,
            max_d           = max_d,
            min_overlap_len = min_length,
        )
    elif 'wall' in cat_lower:
        analyzed = analyze_condition(filtered, category)
    else:
        analyzed = filtered

    if not analyzed:
        cond.AnalysisStatus = 'x'
        MessageBox.Show(
            u"Loc duoc {} elements nhung phan tich theo category '{}' cho ket qua rong.\n"
            u"Kiem tra lai rules / category.".format(len(filtered), category),
            u"Analysis – Khong khop category"
        )
        return

    # 3) Group theo kích thước
    # 3-pre) Pair texts trước khi group (để h được điền vào BeamAxis trước khi tạo group key)
    if 'framing' in cat_lower or 'beam' in cat_lower:
        text_layer_name = ''
        for r in (rules or []):
            rd = r.to_dict() if hasattr(r, 'to_dict') else r
            if rd.get('parameter') == 'Text Layer':
                text_layer_name = (rd.get('value', '') or '').strip().upper()
                break
        if text_layer_name:
            all_texts = vm.get_texts_for_file(filename)
            layer_texts = [(c, x, y, lyr) for c, x, y, lyr in all_texts
                           if lyr.upper() == text_layer_name]
            if layer_texts:
                _pair_texts_with_beams(analyzed, layer_texts)

    groups_data = group_elements_by_label(analyzed)

    # 3b) Remap labels theo Category
    cat_lo = category.lower()
    for g in groups_data:
        shape = g.get('shape', '')
        w_v   = g.get('w', 0)
        h_v   = g.get('h', 0)
        d_v   = g.get('dia', 0)
        if shape == 'REC':
            if 'column' in cat_lo:
                g['label'] = u'CLN: {}x{}'.format(w_v, h_v)
            elif 'framing' in cat_lo or 'beam' in cat_lo:
                g['label'] = u'FRM: {}x{}'.format(w_v, h_v)
            elif 'foundation' in cat_lo or 'footing' in cat_lo:
                g['label'] = u'FDN: {}x{}'.format(w_v, h_v)
            elif 'wall' in cat_lo:
                g['label'] = u'Wall: {}'.format(w_v)   # w = min dim = thickness
        elif shape == 'CIR':
            if 'column' in cat_lo or 'foundation' in cat_lo or 'footing' in cat_lo:
                g['label'] = u'D: {}'.format(d_v)
        elif shape == 'BEA':
            if 'framing' in cat_lo or 'beam' in cat_lo:
                # BeamAxis: w = bề rộng (khoảng cách 2 mép), h = chiều cao (từ text CAD hoặc ?)
                if h_v and h_v > 0:
                    g['label'] = u'FRM: {}x{}'.format(w_v, h_v)
                else:
                    g['label'] = u'FRM: {}x?'.format(w_v)

    # 4) Gán _group_label vào mỗi element (cho canvas coloring)
    for g in groups_data:
        for elem in g['elements']:
            elem._group_label = g['label']
            elem._condition   = cond

    # 5) Tạo CadGroup objects
    cad_groups = []
    for g in groups_data:
        cg = CadGroup(
            label       = g['label'],
            shape       = g['shape'],
            elements    = list(g['elements']),
            w           = g.get('w', 0),
            h           = g.get('h', 0),
            dia         = g.get('dia', 0),
            vm          = vm,
            source_type = 'line' if g['shape'] == 'BEA' else 'point',
            condition   = cond,
        )
        cad_groups.append(cg)

    # 6) Ghi vào condition
    cond.result_elements = analyzed
    cond.cad_groups      = cad_groups
    cond.AnalysisStatus  = 'v'

    # 7) Axis Align: snap locations to nearest 5mm in grid coordinate system
    grid_layer = (getattr(vm, 'CadGridLayer', '') or '').strip().upper()
    axis_status = ''
    if grid_layer and analyzed:
        file_elems = vm.get_elements_for_file(filename)
        grid_elem = None
        for fe in file_elems:
            if (getattr(fe, 'layer', '').upper() == grid_layer and
                    getattr(fe, 'type', '') in ('line', 'polyline') and
                    len(getattr(fe, 'points', [])) >= 2):
                grid_elem = fe
                break
        if grid_elem is not None:
            _, axis_status = align_elements_to_axis(analyzed, grid_elem)
        else:
            axis_status = '!'
    cond.AxisAlignStatus = axis_status

    n_types   = len(cad_groups)
    n_elems   = sum(len(g.elements) for g in cad_groups)
    vm.Status = u"[{}] Analysis xong: {} elements, {} loai.".format(
        cond.ConditionName, n_elems, n_types)

    # Nếu đang preview → refresh
    if cond.PreviewChecked:
        vm.refresh_preview_groups()
        _redraw(window)


# ─────────────────────────────────────────────────────────
#   SELECT LINE IN REVIT (per condition)
# ─────────────────────────────────────────────────────────
def on_select_line_for_condition(cond, window):
    """Raise ExternalEvent chọn đường tham chiếu Revit cho 1 condition cụ thể."""
    if hasattr(window, '_ext_select_line'):
        window._ext_select_line.pending_condition = cond
        window._ext_select_line_event.Raise()
    else:
        MessageBox.Show(u"ExternalEvent chua khoi tao.", u"Loi")


# ─────────────────────────────────────────────────────────
#   SET UP DATA TYPE – DATAGRID SELECTION
# ─────────────────────────────────────────────────────────
_canvas_selecting = False


def on_groups_selection_changed(sender, e):
    """Click dòng Set up data type → highlight group trên Canvas."""
    global _canvas_selecting
    if _canvas_selecting:
        return
    try:
        window = getattr(sender, 'Tag', None)
        if window is None: return
        vm  = window.DataContext
        sel = sender.SelectedItem
        vm.SelectedElement = None
        vm.SelectedGroup   = sel
        _redraw(window)
    except Exception as ex:
        print("on_groups_selection_changed error: {}".format(ex))


# ─────────────────────────────────────────────────────────
#   CANVAS DRAWING
# ─────────────────────────────────────────────────────────
_MARGIN = 20


def _get_bbox(elem_cond_pairs, grid_elems=None):
    all_x, all_y = [], []

    def _add(elem):
        if elem.type in ('polyline', 'line') and getattr(elem, 'points', []):
            for x, y in elem.points:
                all_x.append(x); all_y.append(y)
        elif elem.type in ('beam_line', 'beam_axis'):
            all_x.extend([elem.start[0], elem.end[0]])
            all_y.extend([elem.start[1], elem.end[1]])
        elif elem.type in ('circle', 'arc') and elem.center:
            cx, cy, r = elem.center[0], elem.center[1], elem.radius
            all_x.extend([cx-r, cx+r]); all_y.extend([cy-r, cy+r])

    for (elem, _cond) in (elem_cond_pairs or []):
        _add(elem)
    for elem in (grid_elems or []):
        _add(elem)

    if not all_x: return None
    return (min(all_x), min(all_y), max(all_x), max(all_y))


def _make_transform(bbox, canvas_w, canvas_h, margin=_MARGIN):
    min_x, min_y, max_x, max_y = bbox
    w = max_x - min_x or 1.0
    h = max_y - min_y or 1.0
    uw = canvas_w - 2*margin
    uh = canvas_h - 2*margin
    scale = min(uw/w, uh/h)
    ox = margin + (uw - w*scale)/2.0
    oy = margin + (uh - h*scale)/2.0

    def to_canvas(x, y):
        return (x - min_x)*scale + ox, (max_y - y)*scale + oy

    return to_canvas, scale


def _hex_to_color(hex_str):
    """Chuyển '#RRGGBB' hoặc '#AARRGGBB' sang WPF Color."""
    s = hex_str.lstrip('#')
    if len(s) == 6:
        r, g, b = int(s[0:2],16), int(s[2:4],16), int(s[4:6],16)
        return Color.FromRgb(r, g, b)
    if len(s) == 8:
        a, r, g, b = int(s[0:2],16), int(s[2:4],16), int(s[4:6],16), int(s[6:8],16)
        return Color.FromArgb(a, r, g, b)
    return Colors.Cyan


def _dim_color(hex_str, alpha=70):
    """Version mờ của màu condition."""
    c = _hex_to_color(hex_str)
    return Color.FromArgb(alpha, c.R, c.G, c.B)


def _draw_polyline_on_canvas(canvas, pts, color, thickness=1.5, closed=False, tag=None):
    if closed and len(pts) >= 3:
        pg = Polygon()
        pg.Stroke = SolidColorBrush(color)
        pg.StrokeThickness = thickness
        pg.Fill = SolidColorBrush(Colors.Transparent)
        # Deduplicate last point if same as first
        draw_pts = pts[:-1] if (len(pts) > 1 and
            abs(pts[0][0]-pts[-1][0]) < 0.01 and abs(pts[0][1]-pts[-1][1]) < 0.01) else pts
        for cx, cy in draw_pts:
            pg.Points.Add(Point(cx, cy))
        if tag is not None: pg.Tag = tag
        canvas.Children.Add(pg)
        return pg
    else:
        pl = Polyline()
        pl.Stroke = SolidColorBrush(color)
        pl.StrokeThickness = thickness
        pl.Fill = SolidColorBrush(Colors.Transparent)
        for cx, cy in pts:
            pl.Points.Add(Point(cx, cy))
        if tag is not None: pl.Tag = tag
        canvas.Children.Add(pl)
        return pl


def _draw_line_on_canvas(canvas, x1, y1, x2, y2, color, thickness=1.5, tag=None):
    ln = Line()
    ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
    ln.Stroke = SolidColorBrush(color)
    ln.StrokeThickness = thickness
    if tag is not None: ln.Tag = tag
    canvas.Children.Add(ln)
    return ln


def _draw_circle_on_canvas(canvas, cx, cy, r_px, color, thickness=1.5, tag=None):
    el = Ellipse()
    el.Width, el.Height = r_px*2, r_px*2
    el.Stroke = SolidColorBrush(color)
    el.StrokeThickness = thickness
    el.Fill = SolidColorBrush(Colors.Transparent)
    Canvas.SetLeft(el, cx - r_px)
    Canvas.SetTop(el,  cy - r_px)
    if tag is not None: el.Tag = tag
    canvas.Children.Add(el)
    return el


def _draw_label_on_canvas(canvas, text, cx, cy, color=Colors.White, font_size=9):
    tb = TextBlock()
    tb.Text = text
    tb.Foreground = SolidColorBrush(color)
    tb.FontSize = font_size
    Canvas.SetLeft(tb, cx + 2)
    Canvas.SetTop(tb,  cy - font_size - 2)
    canvas.Children.Add(tb)


class _ShapeTag:
    """Metadata gắn vào WPF Shape.Tag để hỗ trợ hit-test."""
    __slots__ = ('group_label', 'elem', 'condition')
    def __init__(self, group_label, elem, condition):
        self.group_label = group_label
        self.elem      = elem
        self.condition = condition


def _draw_elements_on_canvas(canvas, elem_cond_pairs, grid_elems,
                              canvas_w, canvas_h, selected_label=None, selected_elem=None):
    canvas.Children.Clear()
    if not elem_cond_pairs and not grid_elems:
        return

    bbox = _get_bbox(elem_cond_pairs, grid_elems)
    if not bbox:
        return

    to_canvas, scale = _make_transform(bbox, canvas_w, canvas_h)

    has_elem_sel  = selected_elem is not None
    has_group_sel = selected_label is not None and not has_elem_sel
    has_any_sel   = has_elem_sel or has_group_sel

    def _get_color(elem_obj, lbl, cond):
        base_hex = getattr(cond, 'color', '#00E5FF') if cond else '#00E5FF'
        if not has_any_sel:
            return _hex_to_color(base_hex)
        if has_elem_sel:
            return Colors.White if (elem_obj is selected_elem) else _dim_color(base_hex)
        # group selection
        return _hex_to_color(base_hex) if (lbl == selected_label) else _dim_color(base_hex)

    for (elem, cond) in (elem_cond_pairs or []):
        lbl   = getattr(elem, '_group_label', None)
        color = _get_color(elem, lbl, cond)
        stag  = _ShapeTag(lbl, elem, cond)

        if elem.type == 'polyline' and len(elem.points) >= 2:
            pts = [to_canvas(x, y) for x, y in elem.points]
            is_closed = (len(pts) >= 3 and
                abs(pts[0][0]-pts[-1][0]) < 1.0 and abs(pts[0][1]-pts[-1][1]) < 1.0)
            _draw_polyline_on_canvas(canvas, pts, color, closed=is_closed, tag=stag)

        elif elem.type == 'line' and len(elem.points) == 2:
            c1 = to_canvas(*elem.points[0])
            c2 = to_canvas(*elem.points[1])
            _draw_line_on_canvas(canvas, c1[0], c1[1], c2[0], c2[1], color, tag=stag)

        elif elem.type == 'circle' and elem.center:
            cx, cy = to_canvas(*elem.center)
            _draw_circle_on_canvas(canvas, cx, cy, elem.radius * scale, color, tag=stag)

        elif elem.type in ('beam_line', 'beam_axis'):
            c1 = to_canvas(elem.start[0], elem.start[1])
            c2 = to_canvas(elem.end[0],   elem.end[1])
            _draw_line_on_canvas(canvas, c1[0], c1[1], c2[0], c2[1], color, thickness=2.5, tag=stag)
            # Draw the 2 bounding lines of the beam axis (beam width visualization)
            if elem.type == 'beam_axis':
                half_w = getattr(elem, 'width', 0) / 2.0
                if half_w > 0:
                    # Direction unit vector
                    dx = elem.end[0] - elem.start[0]
                    dy = elem.end[1] - elem.start[1]
                    L  = math.sqrt(dx*dx + dy*dy) or 1.0
                    ux, uy = dx/L, dy/L
                    # Perpendicular
                    vx, vy = -uy, ux
                    bound_color = Color.FromArgb(80, color.R, color.G, color.B)
                    for sign in (1, -1):
                        ox = sign * half_w * vx
                        oy = sign * half_w * vy
                        b1 = to_canvas(elem.start[0]+ox, elem.start[1]+oy)
                        b2 = to_canvas(elem.end[0]+ox,   elem.end[1]+oy)
                        _draw_line_on_canvas(canvas, b1[0], b1[1], b2[0], b2[1],
                                             bound_color, thickness=1.0)

        elif elem.type == 'arc' and elem.center:
            r = elem.radius
            sa, ea = elem.start_angle, elem.end_angle
            if ea < sa: ea += 2*math.pi
            steps = max(16, int((ea-sa)/(math.pi/16)))
            arc_pts = [to_canvas(elem.center[0] + r*math.cos(sa+(ea-sa)*i/steps),
                                  elem.center[1] + r*math.sin(sa+(ea-sa)*i/steps))
                       for i in range(steps+1)]
            _draw_polyline_on_canvas(canvas, arc_pts, color, tag=stag)

    # Vẽ đường tham chiếu CAD (màu vàng)
    for idx, elem in enumerate(grid_elems or []):
        if elem.type in ('line', 'polyline') and len(elem.points) >= 2:
            pts = [to_canvas(x, y) for x, y in elem.points]
            _draw_polyline_on_canvas(canvas, pts, Colors.Yellow, thickness=2.5)
            _draw_label_on_canvas(canvas, 'Ref', pts[0][0], pts[0][1], Colors.Yellow, 11)

    # Mini trục tọa độ
    orig     = to_canvas(bbox[0], bbox[1])
    axis_len = 40
    _draw_line_on_canvas(canvas, orig[0], orig[1], orig[0]+axis_len, orig[1], Colors.Red, 1)
    _draw_label_on_canvas(canvas, 'X', orig[0]+axis_len, orig[1], Colors.Red, 9)
    _draw_line_on_canvas(canvas, orig[0], orig[1], orig[0], orig[1]-axis_len, Colors.LimeGreen, 1)
    _draw_label_on_canvas(canvas, 'Y', orig[0], orig[1]-axis_len-12, Colors.LimeGreen, 9)


# ─────────────────────────────────────────────────────────
#   CANVAS CLICK – hit-test
# ─────────────────────────────────────────────────────────
def on_canvas_click(sender, e):
    try:
        window = sender.Tag
        vm     = window.DataContext
        pt     = e.GetPosition(sender)

        hit_tag = None
        for child in reversed(list(sender.Children)):
            tag = getattr(child, 'Tag', None)
            if not isinstance(tag, _ShapeTag):
                continue
            if _hit_test(child, pt):
                hit_tag = tag
                break

        if hit_tag is None:
            vm.SelectedElement  = None
            vm.SelectedGroup    = None
            # SelectedBeamInfo ở panel cố định – giữ nguyên khi click vùng trống canvas
            _redraw(window)
            return

        vm.SelectedElement = hit_tag.elem
        # Highlight dòng tương ứng trong Set up data type
        global _canvas_selecting
        _canvas_selecting = True
        try:
            for g in vm.PreviewGroups:
                if g.label == hit_tag.group_label:
                    vm._selected_group = g
                    vm.OnPropertyChanged('SelectedGroup')
                    window.DgGroups.ScrollIntoView(g)
                    break
        finally:
            _canvas_selecting = False

        # Nếu click vào beam_axis → hiển thị overlay panel
        elem = hit_tag.elem
        if getattr(elem, 'type', None) == 'beam_axis':
            beam_group = None
            beam_cond  = getattr(hit_tag, 'condition', None)
            if beam_cond is not None:
                for g in beam_cond.cad_groups:
                    if elem in g.elements:
                        beam_group = g
                        break
            if beam_group is not None:
                vm.SelectedBeamInfo = SelectedBeamInfoRow(elem, beam_group, beam_cond, vm)
            else:
                vm.SelectedBeamInfo = None
        else:
            vm.SelectedBeamInfo = None

        _redraw(window)
    except Exception as ex:
        print("on_canvas_click error: {}".format(ex))


def _hit_test(shape, pt):
    try:
        from System.Windows import Rect
        if isinstance(shape, Line):
            x1, y1, x2, y2 = shape.X1, shape.Y1, shape.X2, shape.Y2
            tol = max(float(shape.StrokeThickness)*2.0, 8.0)
            dx, dy = x2-x1, y2-y1
            len_sq = dx*dx + dy*dy
            if len_sq < 1e-6:
                return math.sqrt((pt.X-x1)**2+(pt.Y-y1)**2) <= tol
            t  = max(0.0, min(1.0, ((pt.X-x1)*dx+(pt.Y-y1)*dy)/len_sq))
            nx, ny = x1+t*dx-pt.X, y1+t*dy-pt.Y
            return math.sqrt(nx*nx+ny*ny) <= tol
        rg = getattr(shape, 'RenderedGeometry', None)
        if rg is not None:
            b   = rg.Bounds
            pad = max(float(getattr(shape, 'StrokeThickness', 1.5))+4.0, 6.0)
            return Rect(b.X-pad, b.Y-pad, b.Width+pad*2, b.Height+pad*2).Contains(pt)
        left = Canvas.GetLeft(shape)
        top  = Canvas.GetTop(shape)
        w    = float(getattr(shape, 'ActualWidth',  0) or 8)
        h    = float(getattr(shape, 'ActualHeight', 0) or 8)
        if w < 8: w = 8
        if h < 8: h = 8
        if left != left or top != top: return False   # NaN guard
        return Rect(left-4, top-4, w+8, h+8).Contains(pt)
    except Exception:
        return False


# ─────────────────────────────────────────────────────────
#   REDRAW
# ─────────────────────────────────────────────────────────
def _redraw(window):
    canvas = window.DrawingCanvas
    w = canvas.ActualWidth
    h = canvas.ActualHeight
    if w < 10 or h < 10:
        return
    vm = window.DataContext

    sel_elem  = getattr(vm, 'SelectedElement', None)
    sel_label = None
    if sel_elem is None:
        sg = getattr(vm, 'SelectedGroup', None)
        if sg is not None:
            sel_label = getattr(sg, 'label', None)

    pairs = vm.get_previewed_elements_with_condition()
    _draw_elements_on_canvas(
        canvas, pairs,
        getattr(vm, 'cad_grid_elements', []),
        w, h,
        selected_label=sel_label,
        selected_elem=sel_elem,
    )


def on_canvas_size_changed(sender, e):
    try:
        _redraw(sender.Tag)
    except Exception:
        pass
