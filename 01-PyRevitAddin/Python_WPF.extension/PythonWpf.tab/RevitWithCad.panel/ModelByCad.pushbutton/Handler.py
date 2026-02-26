# -*- coding: utf-8 -*-
"""
Handler.py - Gắn sự kiện cho cửa sổ ModelByCad.
"""
import math
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows.Shapes import Polyline, Line, Ellipse, Polygon
from System.Windows.Controls import Canvas, TextBlock
from System.Windows.Media import SolidColorBrush, Colors, Color
from System.Windows import Point, Thickness
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult as WFDialogResult

from CadUtils import (
    get_acad_doc, select_elements_in_cad, select_grid_in_cad,
    merge_lines_to_closed_polylines, group_elements_by_label,
    select_beam_elements_in_cad, group_beam_pairs_by_label
)


# ============================================================
#   REVIT EXTERNAL EVENT HANDLERS (cho cửa sổ modeless)
# ============================================================
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


# ============================================================
#   GEOMETRY HELPERS – CAD → REVIT TRANSFORM
# ============================================================
_MM_TO_FEET = 1.0 / 304.8


def _get_cad_ref(cad_elem):
    """
    Lấy (start_x_mm, start_y_mm, dir_x, dir_y) từ CadElement tham chiếu.
    points[0] = StartPoint chọn trong CAD, points[-1] = EndPoint.
    """
    s = cad_elem.points[0]
    e = cad_elem.points[-1]
    dx = e[0] - s[0]
    dy = e[1] - s[1]
    length = math.sqrt(dx * dx + dy * dy) or 1.0
    return (s[0], s[1], dx / length, dy / length)


def _get_revit_ref(revit_elem):
    """
    Lấy (start_x_ft, start_y_ft, dir_x, dir_y) từ Revit element.
    Hỗ trợ Grid (có .Curve) và CurveElement / ModelLine / DetailLine (có .Location.Curve).
    """
    try:
        # Grid
        curve = revit_elem.Curve
    except Exception:
        # ModelLine, DetailLine, CurveElement…
        curve = revit_elem.Location.Curve
    start = curve.GetEndPoint(0)
    d     = curve.Direction
    return (start.X, start.Y, d.X, d.Y)


def _build_transform(cad_elem, revit_elem):
    """
    Tính bộ 3 (s_cad, s_revit_ft, theta) từ cặp đường tham chiếu.
    - s_cad     : (x_mm, y_mm)  – StartPoint CAD   = origin không gian CAD
    - s_revit_ft: (x_ft, y_ft)  – StartPoint Revit  = origin không gian Revit
    - theta     : góc xoay (radian) = angle(Revit) − angle(CAD)
    """
    sx_c, sy_c, dx_c, dy_c = _get_cad_ref(cad_elem)
    sx_r, sy_r, dx_r, dy_r = _get_revit_ref(revit_elem)
    theta = math.atan2(dy_r, dx_r) - math.atan2(dy_c, dx_c)
    return (sx_c, sy_c), (sx_r, sy_r), theta


def _rotate2d(dx, dy, theta):
    """Xoay vector (dx, dy) quanh gốc một góc theta (radian)."""
    cos_t = math.cos(theta)
    sin_t = math.sin(theta)
    return (dx * cos_t - dy * sin_t,
            dx * sin_t + dy * cos_t)


def _transform_to_revit(cx_mm, cy_mm, s_cad, s_revit_ft, theta):
    """
    Chuyển tọa độ tâm CAD (mm) sang XYZ Revit (feet).
    s_cad      : (x_mm, y_mm)  – StartPoint đường tham chiếu CAD
    s_revit_ft : (x_ft, y_ft)  – StartPoint đường tham chiếu Revit
    theta      : góc xoay (radian) = angle(Revit ref) − angle(CAD ref)
    """
    dx_mm = cx_mm - s_cad[0]
    dy_mm = cy_mm - s_cad[1]
    dx_rot, dy_rot = _rotate2d(dx_mm, dy_mm, theta)
    x_ft = s_revit_ft[0] + dx_rot * _MM_TO_FEET
    y_ft = s_revit_ft[1] + dy_rot * _MM_TO_FEET
    return XYZ(x_ft, y_ft, 0.0)


def _get_element_center(elem):
    """
    Tính tọa độ tâm của 1 CadElement / _MergedPolyline.
    Trả về (cx_mm, cy_mm) hoặc None.
    """
    if elem.type == 'circle' and elem.center:
        return elem.center
    if elem.type in ('polyline', 'line') and elem.points:
        xs = [p[0] for p in elem.points]
        ys = [p[1] for p in elem.points]
        # Dùng bounding-box centroid (chính xác hơn mean cho hình chữ nhật)
        return ((min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0)
    if elem.type == 'arc' and elem.center:
        return elem.center
    return None


_CAT_TO_STRUCTURAL = {
    'structural columns':    None,   # StructuralType.Column     – set sau khi import
    'structural foundations': None,  # StructuralType.Footing
    'structural framing':    None,   # StructuralType.Beam
}


def _init_structural_types():
    """Điền giá trị enum sau khi import thành công."""
    if not _REVIT_AVAILABLE:
        return
    _CAT_TO_STRUCTURAL['structural columns']     = StructuralType.Column
    _CAT_TO_STRUCTURAL['structural foundations'] = StructuralType.Footing
    _CAT_TO_STRUCTURAL['structural framing']     = StructuralType.Beam


if _REVIT_AVAILABLE:
    _init_structural_types()


def _category_to_structural_type(cat_name):
    """Map tên Category → StructuralType enum. Mặc định NonStructural."""
    return _CAT_TO_STRUCTURAL.get(cat_name.lower(), StructuralType.NonStructural)


def _resolve_symbols(doc):
    """
    Trả về dict {'FamilyName : TypeName': FamilySymbol} cho toàn bộ doc.
    Revit 2024: dùng FamilyName property + BuiltInParameter fallback thay vì
    sym.Family.Name / sym.Name trực tiếp để tránh crash.
    """
    result = {}
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            # Family name
            fam_name = ''
            try:
                fam_name = sym.FamilyName or ''
            except Exception:
                pass
            if not fam_name:
                try:
                    if sym.Family is not None:
                        fam_name = sym.Family.Name or ''
                except Exception:
                    pass
            if not fam_name:
                continue

            # Type name
            sym_name = ''
            try:
                sym_name = sym.Name or ''
            except Exception:
                pass
            if not sym_name:
                try:
                    p = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if p is not None:
                        sym_name = p.AsString() or ''
                except Exception:
                    pass
            if not sym_name:
                try:
                    p = sym.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    if p is not None:
                        sym_name = p.AsString() or ''
                except Exception:
                    pass
            if not sym_name:
                continue

            key = "{} : {}".format(fam_name, sym_name)
            result[key] = sym
        except Exception:
            continue
    return result


def _resolve_levels(doc):
    """Trả về dict {'Level name': Level}."""
    result = {}
    for lv in FilteredElementCollector(doc).OfClass(Level):
        try:
            result[lv.Name] = lv
        except Exception:
            continue
    return result


class SelectRevitGridHandler(IExternalEventHandler if _REVIT_AVAILABLE else object):
    """
    IExternalEventHandler: chọn 1 đường tham chiếu trong Revit (Grid hoặc Line).
    StartPoint của đường = điểm gốc transform; phương chiều = hướng tham chiếu.
    """
    def __init__(self, vm):
        self.vm = vm

    def Execute(self, uiapp):
        try:
            uidoc = uiapp.ActiveUIDocument
            doc   = uidoc.Document
            ref   = uidoc.Selection.PickObject(
                ObjectType.Element,
                "Chon 1 duong tham chieu trong Revit (Grid hoac Line)"
            )
            elem = doc.GetElement(ref.ElementId)
            # Kiểm tra element có curve hay không
            try:
                _ = _get_revit_ref(elem)
            except Exception:
                print("SelectRevitGridHandler: element khong co Curve, thu chon lai.")
                return
            self.vm.revit_grids = [elem]
            self.vm.Status = "Da chon duong tham chieu Revit."
        except Exception as ex:
            print("SelectRevitGridHandler error: {}".format(ex))

    def GetName(self):
        return "SelectRevitGrid"


class CreateModelHandler(IExternalEventHandler if _REVIT_AVAILABLE else object):
    """
    IExternalEventHandler: tạo FamilyInstance trong Revit từ dữ liệu CAD.
    Luồng:
      1. Tính giao điểm trục CAD + Revit  →  điểm neo + góc xoay
      2. Với mỗi CadGroup đã điền đủ category/familytype/level:
           resolve FamilySymbol, Level, StructuralType
           tính tâm từng element  →  XYZ Revit
           doc.Create.NewFamilyInstance(...)
    """
    def __init__(self, vm):
        self.vm = vm

    def Execute(self, uiapp):
        vm  = self.vm
        doc = uiapp.ActiveUIDocument.Document

        # ── Kiểm tra dữ liệu đầu vào ────────────────────────────────
        if len(vm.cad_grid_elements) < 1 or len(vm.revit_grids) < 1:
            MessageBox.Show(
                "Chua chon duong tham chieu!\n"
                "Vui long:\n"
                "  1. Nhan 'Select grid in Cad' de chon duong ref trong CAD\n"
                "  2. Nhan 'Select grid in Revit' de chon duong ref trong Revit",
                "Thieu duong tham chieu"
            )
            return

        ready_groups = [g for g in vm.CadGroups if g.is_ready()]
        if not ready_groups:
            MessageBox.Show(
                "Chua co nhom nao dien du thong tin.\n"
                "Vui long dien Category, Family Type va Base Level cho it nhat 1 hang trong bang.",
                "Thieu thong so"
            )
            return

        # ── Tính transform ───────────────────────────────────────────
        try:
            s_cad, s_revit_ft, theta = _build_transform(
                vm.cad_grid_elements[0], vm.revit_grids[0]
            )
        except Exception as ex:
            print("CreateModel: Loi tinh transform: {}".format(ex))
            return

        # ── Resolve lookup tables ────────────────────────────────────
        symbol_map = _resolve_symbols(doc)
        level_map  = _resolve_levels(doc)

        # ── Tạo FamilyInstance trong Transaction ─────────────────────
        created = 0
        skipped = 0
        errors  = []

        t = Transaction(doc, "Create Model from CAD")
        t.Start()
        try:
            for group in ready_groups:
                # Resolve FamilySymbol
                sym_key = group.FamilyType
                symbol  = symbol_map.get(sym_key)
                if symbol is None:
                    skipped += 1
                    errors.append("Khong tim thay FamilySymbol: '{}'\n  (Keys mau: {})".format(
                        sym_key, list(symbol_map.keys())[:3]))
                    continue

                # Activate symbol nếu chưa active
                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()

                # Resolve Level
                level = level_map.get(group.BaseLevel)
                if level is None:
                    skipped += 1
                    errors.append("Khong tim thay Level: '{}'".format(group.BaseLevel))
                    continue

                # Resolve StructuralType
                struct_type = _category_to_structural_type(group.Category)

                # ── Dầm (source_type='line') → dùng curve overload ────────────
                if group.is_beam():
                    for pair in group.elements:
                        try:
                            xyz_s = _transform_to_revit(
                                pair.start[0], pair.start[1], s_cad, s_revit_ft, theta
                            )
                            xyz_e = _transform_to_revit(
                                pair.end[0], pair.end[1],   s_cad, s_revit_ft, theta
                            )
                            curve = RvtLine.CreateBound(xyz_s, xyz_e)
                            inst = doc.Create.NewFamilyInstance(
                                curve, symbol, level, StructuralType.Beam
                            )
                            # ① Y-justification = Left (= 1)
                            try:
                                param = inst.get_Parameter(BuiltInParameter.Y_JUSTIFICATION)
                                if param is not None and not param.IsReadOnly:
                                    param.Set(3)
                            except Exception:
                                pass
                            created += 1
                        except Exception as ex_beam:
                            skipped += 1
                            errors.append(str(ex_beam))
                    continue   # next group

                # ── Cột / Móng (source_type='point') → dùng XYZ overload ────
                # Tạo instance cho từng element trong nhóm
                for elem in group.elements:
                    center = _get_element_center(elem)
                    if center is None:
                        skipped += 1
                        continue
                    try:
                        xyz = _transform_to_revit(
                            center[0], center[1], s_cad, s_revit_ft, theta
                        )
                        doc.Create.NewFamilyInstance(
                            xyz, symbol, level, struct_type
                        )
                        created += 1
                    except Exception as ex_inst:
                        skipped += 1
                        errors.append(str(ex_inst))

            t.Commit()
        except Exception as ex_t:
            t.RollBack()
            print("CreateModel: Transaction loi, da rollback: {}".format(ex_t))
            return

        # ── Báo cáo ──────────────────────────────────────────────────
        vm.Status = "Tao xong: {} instance. Bo qua: {}.".format(created, skipped)
        print("CreateModel: {} created, {} skipped.".format(created, skipped))
        for msg in errors:
            print("  [WARN] {}".format(msg))
        if errors:
            MessageBox.Show(
                "Tao {} instance.\nBo qua {} vi:\n{}".format(
                    created, skipped, "\n".join(errors[:5])),
                "Ket qua tao model"
            )
        else:
            MessageBox.Show(
                "Da tao thanh cong {} instance!".format(created),
                "Hoan thanh"
            )

    def GetName(self):
        return "CreateModel"


# ============================================================
#   BIND HANDLERS
# ============================================================
def bind_handlers(window):
    """Gắn tất cả event handlers vào window."""
    window.BtnSelectElements.Tag = window
    window.BtnSelectElements.Click += on_select_elements_in_cad

    window.BtnSelectBeams.Tag = window
    window.BtnSelectBeams.Click += on_select_beam_elements_in_cad

    window.BtnSelectGridCad.Tag = window
    window.BtnSelectGridCad.Click += on_select_grid_in_cad

    window.BtnSelectGridRevit.Tag = window
    window.BtnSelectGridRevit.Click += on_select_grid_in_revit

    window.BtnCreateModel.Tag = window
    window.BtnCreateModel.Click += on_create_model

    window.BtnDelete.Tag = window
    window.BtnDelete.Click += on_delete

    window.BtnRefresh.Tag = window
    window.BtnRefresh.Click += on_refresh

    # Canvas click for shape → individual element selection
    window.DrawingCanvas.Tag = window
    window.DrawingCanvas.MouseLeftButtonDown += on_canvas_click
    window.DrawingCanvas.SizeChanged += on_canvas_size_changed

    # DataGrid: Tag=window so selection handler can reach the window
    window.DgGroups.Tag = window
    window.DgGroups.SelectionChanged += on_datagrid_selection_changed


# ============================================================
#   COLOUR HELPERS
# ============================================================
_COLOR_NORMAL   = Colors.Cyan          # shape unselected
_COLOR_SELECTED = Colors.White         # shape in selected group
_COLOR_DIM_A    = 80                   # alpha for dimmed shapes

def _dim_color():
    """Cyan с пониженной прозрачностью для dimmed shapes."""
    return Color.FromArgb(_COLOR_DIM_A, 0, 255, 255)


# ============================================================
#   SHAPE TAG  –  gắn vào mọi WPF shape để biết element + nhóm
# ============================================================
class _ShapeTag(object):
    """Metadata gắn vào WPF Shape.Tag để hỗ trợ canvas-click selection."""
    __slots__ = ('group_label', 'elem')
    def __init__(self, group_label, elem):
        self.group_label = group_label   # str: group.label
        self.elem        = elem          # CadElement / CadBeamPair object


# ============================================================
#   PREVIEW DRAWING HELPERS
# ============================================================
_MARGIN = 20  # px


def _get_bbox(elements, grid_elements=None):
    """Tính bounding box của tất cả elements + grid elements."""
    all_x = []
    all_y = []

    def _add_elem(elem):
        if elem.type in ('polyline', 'line') and elem.points:
            for x, y in elem.points:
                all_x.append(x)
                all_y.append(y)
        elif elem.type == 'beam_line':
            all_x.extend([elem.start[0], elem.end[0]])
            all_y.extend([elem.start[1], elem.end[1]])
        elif elem.type in ('circle', 'arc') and elem.center:
            cx, cy = elem.center
            r = elem.radius
            all_x.extend([cx - r, cx + r])
            all_y.extend([cy - r, cy + r])

    for e in (elements or []):
        _add_elem(e)
    for e in (grid_elements or []):
        _add_elem(e)

    if not all_x or not all_y:
        return None
    return (min(all_x), min(all_y), max(all_x), max(all_y))


def _make_transform(bbox, canvas_w, canvas_h, margin=_MARGIN):
    """
    Tạo hàm chuyển đổi tọa độ CAD → Canvas.
    Flip trục Y (CAD tăng lên, Canvas tăng xuống).
    """
    min_x, min_y, max_x, max_y = bbox
    w = max_x - min_x or 1.0
    h = max_y - min_y or 1.0

    usable_w = canvas_w - 2 * margin
    usable_h = canvas_h - 2 * margin

    scale = min(usable_w / w, usable_h / h)

    offset_x = margin + (usable_w - w * scale) / 2.0
    offset_y = margin + (usable_h - h * scale) / 2.0

    def to_canvas(x, y):
        cx = (x - min_x) * scale + offset_x
        cy = (max_y - y) * scale + offset_y
        return (cx, cy)

    return to_canvas, scale


def _draw_polyline_shape(canvas, pts_canvas, color, thickness=1.5, closed=False, tag=None):
    """Vẽ Polyline (hở) hoặc Polygon (kín) WPF."""
    if closed and len(pts_canvas) >= 3:
        pg = Polygon()
        pg.Stroke = SolidColorBrush(color)
        pg.StrokeThickness = thickness
        pg.Fill = SolidColorBrush(Colors.Transparent)
        pts_dedup = pts_canvas[:-1] if (
            len(pts_canvas) > 1 and
            abs(pts_canvas[0][0] - pts_canvas[-1][0]) < 0.01 and
            abs(pts_canvas[0][1] - pts_canvas[-1][1]) < 0.01
        ) else pts_canvas
        for cx, cy in pts_dedup:
            pg.Points.Add(Point(cx, cy))
        if tag is not None:
            pg.Tag = tag
        canvas.Children.Add(pg)
        return pg
    else:
        pl = Polyline()
        pl.Stroke = SolidColorBrush(color)
        pl.StrokeThickness = thickness
        pl.Fill = SolidColorBrush(Colors.Transparent)
        for cx, cy in pts_canvas:
            pl.Points.Add(Point(cx, cy))
        if tag is not None:
            pl.Tag = tag
        canvas.Children.Add(pl)
        return pl


def _draw_line_shape(canvas, x1, y1, x2, y2, color, thickness=1.5, tag=None):
    """Vẽ một Line WPF."""
    ln = Line()
    ln.X1 = x1; ln.Y1 = y1
    ln.X2 = x2; ln.Y2 = y2
    ln.Stroke = SolidColorBrush(color)
    ln.StrokeThickness = thickness
    if tag is not None:
        ln.Tag = tag
    canvas.Children.Add(ln)
    return ln


def _draw_circle_shape(canvas, cx, cy, r_px, color, thickness=1.5, tag=None):
    """Vẽ Ellipse (circle) WPF."""
    el = Ellipse()
    el.Width = r_px * 2
    el.Height = r_px * 2
    el.Stroke = SolidColorBrush(color)
    el.StrokeThickness = thickness
    el.Fill = SolidColorBrush(Colors.Transparent)
    Canvas.SetLeft(el, cx - r_px)
    Canvas.SetTop(el, cy - r_px)
    if tag is not None:
        el.Tag = tag
    canvas.Children.Add(el)
    return el


def _draw_label(canvas, text, cx, cy, color=Colors.White, font_size=10):
    tb = TextBlock()
    tb.Text = text
    tb.Foreground = SolidColorBrush(color)
    tb.FontSize = font_size
    Canvas.SetLeft(tb, cx + 3)
    Canvas.SetTop(tb, cy - font_size - 2)
    canvas.Children.Add(tb)


def _draw_elements_on_canvas(canvas, elements, grid_elements, canvas_w, canvas_h,
                              selected_label=None, selected_elem=None):
    """
    Xóa canvas và vẽ lại tất cả elements + grid.
    - selected_label : label nhóm (DataGrid selection) – tô sáng toàn nhóm
    - selected_elem  : element đơn lẻ (canvas click) – tô sáng 1 shape, ghi đè selected_label
    Trả về dict: {label: [shape_objects]}
    """
    canvas.Children.Clear()

    if not elements and not grid_elements:
        return {}

    bbox = _get_bbox(elements, grid_elements)
    if not bbox:
        return {}

    to_canvas, scale = _make_transform(bbox, canvas_w, canvas_h)

    # --- chế độ highlight ---
    has_elem_sel  = selected_elem is not None
    has_group_sel = (selected_label is not None) and not has_elem_sel
    has_any_sel   = has_elem_sel or has_group_sel

    def _color_for(lbl, elem_obj, default=_COLOR_NORMAL):
        """Màu shape theo selection state. default – màu khi không có selection."""
        if not has_any_sel:
            return default
        if has_elem_sel:
            return _COLOR_SELECTED if (elem_obj is selected_elem) else _dim_color()
        # group selection
        return _COLOR_SELECTED if lbl == selected_label else _dim_color()

    label_to_shapes = {}   # {label: [shape]}

    for elem in (elements or []):
        lbl   = getattr(elem, '_group_label', None)
        shape = None

        if elem.type == 'polyline' and len(elem.points) >= 2:
            color = _color_for(lbl, elem)
            pts = [to_canvas(x, y) for x, y in elem.points]
            is_closed = (
                len(pts) >= 3 and
                abs(pts[0][0] - pts[-1][0]) < 1.0 and
                abs(pts[0][1] - pts[-1][1]) < 1.0
            )
            shape = _draw_polyline_shape(canvas, pts, color, closed=is_closed,
                                         tag=_ShapeTag(lbl, elem))

        elif elem.type == 'line' and len(elem.points) == 2:
            color = _color_for(lbl, elem)
            (x1, y1), (x2, y2) = elem.points
            c1 = to_canvas(x1, y1)
            c2 = to_canvas(x2, y2)
            shape = _draw_line_shape(canvas, c1[0], c1[1], c2[0], c2[1], color,
                                     tag=_ShapeTag(lbl, elem))

        elif elem.type == 'circle' and elem.center:
            color = _color_for(lbl, elem)
            cx, cy = to_canvas(*elem.center)
            r_px = elem.radius * scale
            shape = _draw_circle_shape(canvas, cx, cy, r_px, color,
                                       tag=_ShapeTag(lbl, elem))

        elif elem.type == 'beam_line':
            # default màu dầm = YellowGreen; khi có selection dùng _color_for
            beam_color = _color_for(lbl, elem, default=Colors.YellowGreen)
            c1 = to_canvas(elem.start[0], elem.start[1])
            c2 = to_canvas(elem.end[0],   elem.end[1])
            shape = _draw_line_shape(canvas, c1[0], c1[1], c2[0], c2[1],
                                     beam_color, thickness=2.5, tag=_ShapeTag(lbl, elem))

        elif elem.type == 'arc' and elem.center:
            color = _color_for(lbl, elem)
            cx_w, cy_w = elem.center
            r = elem.radius
            sa, ea = elem.start_angle, elem.end_angle
            if ea < sa:
                ea += 2 * math.pi
            steps = max(16, int((ea - sa) / (math.pi / 16)))
            arc_pts = []
            for i in range(steps + 1):
                a = sa + (ea - sa) * i / steps
                arc_pts.append(to_canvas(cx_w + r * math.cos(a), cy_w + r * math.sin(a)))
            shape = _draw_polyline_shape(canvas, arc_pts, color, tag=_ShapeTag(lbl, elem))

        if shape is not None and lbl is not None:
            label_to_shapes.setdefault(lbl, []).append(shape)

    # Vẽ grid axes (đường tham chiếu, màu vàng, dày hơn)
    axis_labels_txt = ['Ref']
    for idx, elem in enumerate(grid_elements or []):
        color = Colors.Yellow
        label_txt = axis_labels_txt[idx] if idx < len(axis_labels_txt) else str(idx)
        if elem.type in ('line', 'polyline') and len(elem.points) >= 2:
            pts = [to_canvas(x, y) for x, y in elem.points]
            _draw_polyline_shape(canvas, pts, color, thickness=2.5)
            _draw_label(canvas, label_txt, pts[0][0], pts[0][1], color=color, font_size=12)

    # Vẽ mini coordinate axes ở góc dưới trái
    orig = to_canvas(bbox[0], bbox[1])
    axis_len = 40
    _draw_line_shape(canvas, orig[0], orig[1], orig[0] + axis_len, orig[1], Colors.Red, 1)
    _draw_label(canvas, 'X', orig[0] + axis_len, orig[1], Colors.Red, 9)
    _draw_line_shape(canvas, orig[0], orig[1], orig[0], orig[1] - axis_len, Colors.LimeGreen, 1)
    _draw_label(canvas, 'Y', orig[0], orig[1] - axis_len - 12, Colors.LimeGreen, 9)

    return label_to_shapes


# ============================================================
#   CANVAS CLICK HANDLER
# ============================================================
# Flag: set True while on_canvas_click đang cập nhật DataGrid selection
# → ngăn on_datagrid_selection_changed xóa SelectedElement
_canvas_selecting = False


def on_canvas_click(sender, e):
    """
    Click trên Canvas → chọn individual element;
    DataGrid row tương ứng được highlight theo nhóm.
    """
    try:
        window = sender.Tag
        vm = window.DataContext
        if not vm or not vm.CadGroups.Count:
            return

        pt = e.GetPosition(sender)

        hit_tag = None
        children = list(sender.Children)
        for child in reversed(children):
            tag = getattr(child, 'Tag', None)
            if not isinstance(tag, _ShapeTag):
                continue
            if _hit_test(child, pt):
                hit_tag = tag
                break

        if hit_tag is None:
            # Click vào vùng trống → bỏ selection
            vm.SelectedElement = None
            vm.SelectedGroup   = None
            _redraw(window)
            return

        # Chọn individual element (không clear SelectedGroup)
        vm.SelectedElement = hit_tag.elem
        # Đồng thời highlight nhóm trong DataGrid
        # Dùng flag để ngăn on_datagrid_selection_changed xóa SelectedElement
        global _canvas_selecting
        _canvas_selecting = True
        try:
            for group in vm.CadGroups:
                if group.label == hit_tag.group_label:
                    vm._selected_group = group
                    vm.OnPropertyChanged('SelectedGroup')
                    window.DgGroups.ScrollIntoView(group)
                    break
        finally:
            _canvas_selecting = False
        _redraw(window)
    except Exception as ex:
        print("on_canvas_click error: {}".format(ex))


def _hit_test(shape, pt):
    """
    Hit-test điểm pt trên WPF shape.
    - Polygon / Polyline: dùng RenderedGeometry.Bounds
    - Line (X1/Y1/X2/Y2): khoảng cách từ điểm đến đoạn thẳng
    - Ellipse: dùng Canvas.GetLeft/Top + Width/Height
    """
    try:
        from System.Windows import Rect
        # ── WPF Line: hit test bằng khoảng cách từ pt đến segment ──
        if isinstance(shape, Line):
            x1, y1 = shape.X1, shape.Y1
            x2, y2 = shape.X2, shape.Y2
            tolerance = max(float(shape.StrokeThickness) * 2.0, 8.0)
            dx = x2 - x1
            dy = y2 - y1
            len_sq = dx * dx + dy * dy
            if len_sq < 1e-6:
                # Đoạn thẳng suy biến – kiểm tra khoảng cách tới điểm
                dist = math.sqrt((pt.X - x1) ** 2 + (pt.Y - y1) ** 2)
                return dist <= tolerance
            t = max(0.0, min(1.0, ((pt.X - x1) * dx + (pt.Y - y1) * dy) / len_sq))
            nx = x1 + t * dx - pt.X
            ny = y1 + t * dy - pt.Y
            return math.sqrt(nx * nx + ny * ny) <= tolerance

        # ── Polygon / Polyline: dùng RenderedGeometry ──
        rg = getattr(shape, 'RenderedGeometry', None)
        if rg is not None:
            bounds = rg.Bounds
            # Phồng nhẹ để dễ click vào cạnh polyline
            padding = max(float(getattr(shape, 'StrokeThickness', 1.5)) + 4.0, 6.0)
            expanded = Rect(bounds.X - padding, bounds.Y - padding,
                            bounds.Width + padding * 2, bounds.Height + padding * 2)
            return expanded.Contains(pt)

        # ── Ellipse / fallback: Canvas.SetLeft/Top + Width/Height ──
        left = Canvas.GetLeft(shape)
        top  = Canvas.GetTop(shape)
        w    = float(getattr(shape, 'ActualWidth',  0) or 0)
        h    = float(getattr(shape, 'ActualHeight', 0) or 0)
        if w < 8: w = 8
        if h < 8: h = 8
        if left != left or top != top:   # NaN guard
            return False
        return Rect(left - 4, top - 4, w + 8, h + 8).Contains(pt)
    except Exception:
        return False


# ============================================================
#   DATAGRID SELECTION → CANVAS HIGHLIGHT
# ============================================================
def on_datagrid_selection_changed(sender, e):
    """Hàng DataGrid thay đổi → vẽ lại canvas với highlight toàn nhóm."""
    try:
        # Nếu đang xử lý canvas click thì bỏ qua – tránh xóa SelectedElement
        if _canvas_selecting:
            return
        window = getattr(sender, 'Tag', None)
        if window is None:
            return
        vm = window.DataContext
        if vm is None:
            return
        sel = sender.SelectedItem
        vm.SelectedElement = None   # DataGrid selection → xóa individual elem selection
        vm.SelectedGroup   = sel
        _redraw(window)
    except Exception as ex:
        print("on_datagrid_selection_changed error: {}".format(ex))


# ============================================================
#   BUTTON HANDLERS
# ============================================================
def on_delete(sender, e):
    """
    Xóa element đang chọn (canvas click) hoặc nhóm đang chọn (DataGrid row).
    """
    window = sender.Tag
    vm = window.DataContext

    sel_elem  = getattr(vm, 'SelectedElement', None)
    sel_group = getattr(vm, 'SelectedGroup',   None)

    if sel_elem is not None:
        # Xóa 1 element đơn lẻ
        vm.remove_element(sel_elem)
        vm.SelectedElement = None
        vm.SelectedGroup   = None
        vm.Status = "Da xoa 1 element. Con lai {} nhom.".format(vm.CadGroups.Count)
    elif sel_group is not None:
        # Xóa cả nhóm
        n = len(sel_group.elements)
        vm.remove_group(sel_group)
        vm.SelectedElement = None
        vm.SelectedGroup   = None
        vm.Status = "Da xoa nhom '{}' ({} elements). Con lai {} nhom.".format(
            sel_group.label, n, vm.CadGroups.Count
        )
    else:
        MessageBox.Show("Chua chon element hay nhom nao de xoa.\n"
                        "- Click vao shape trong Preview de chon element don le.\n"
                        "- Click vao hang trong DataGrid de chon ca nhom.",
                        "Xoa")
        return

    _redraw(window)


def on_refresh(sender, e):
    """Xóa toàn bộ dữ liệu CAD + DataGrid."""
    res = MessageBox.Show(
        "Xóa toàn bộ dữ liệu (Preview + DataGrid)?",
        "Refresh",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question
    )
    if res != WFDialogResult.Yes:
        return
    window = sender.Tag
    vm = window.DataContext
    vm.clear_all()
    vm.Status = "Da xoa toan bo du lieu."
    _redraw(window)


def on_select_beam_elements_in_cad(sender, e):
    """Chọn dầm trong CAD (line + text pairs), phân tích nhóm, cập nhật DataGrid."""
    window = sender.Tag
    vm = window.DataContext

    doc = get_acad_doc()
    if not doc:
        MessageBox.Show("Khong the ket noi AutoCAD. Hay mo AutoCAD truoc.", "Loi ket noi")
        return

    beam_pairs = select_beam_elements_in_cad(doc)
    if not beam_pairs:
        MessageBox.Show("Khong co cap Line+Text nao duoc chon (hoac khong phan tich duoc kich thuoc WxH).",
                        "Chon Elements (Beam)")
        return

    # Phân nhóm theo label "BEA: WxH"
    groups_data = group_beam_pairs_by_label(beam_pairs)

    # Gán _group_label vào từng CadBeamPair
    for g in groups_data:
        for pair in g['elements']:
            pair._group_label = g['label']

    # Bổ sung vào ViewModel (APPEND, không xóa dữ liệu cũ)
    vm.beam_elements.extend(beam_pairs)
    vm.merge_into_beam_groups(groups_data)

    total_beams  = len(vm.beam_elements)
    total_groups = sum(1 for g in vm.CadGroups if g.is_beam())
    vm.Status = "Them {} dam moi → tong {} dam, {} nhom BEA.".format(
        len(beam_pairs), total_beams, total_groups
    )

    _redraw(window)
    MessageBox.Show(
        "Da them {} dam tu AutoCAD.\nTong cong: {} dam, {} nhom.".format(
            len(beam_pairs), total_beams, total_groups
        ),
        "Hoan thanh"
    )


def on_select_elements_in_cad(sender, e):
    """Chọn elements trong CAD, gom line → polyline kín, phân tích nhóm, cập nhật DataGrid."""
    window = sender.Tag
    vm = window.DataContext

    doc = get_acad_doc()
    if not doc:
        MessageBox.Show("Khong the ket noi AutoCAD. Hay mo AutoCAD truoc.", "Loi ket noi")
        return

    raw_elements = select_elements_in_cad(doc)
    if not raw_elements:
        MessageBox.Show("Khong co element nao duoc chon.", "Chon Elements")
        return

    # 1) Gom các line/polyline thành closed polylines
    elements = merge_lines_to_closed_polylines(raw_elements)

    # 2) Phân tích → nhóm theo kích thước
    groups_data = group_elements_by_label(elements)

    # 3) Gán _group_label vào từng element để canvas biết màu
    for g in groups_data:
        for elem in g['elements']:
            elem._group_label = g['label']

    # 4) Bổ sung vào ViewModel (APPEND, không xóa dữ liệu cũ)
    vm.cad_elements.extend(elements)
    vm.merge_into_groups(groups_data)

    total_elems  = len(vm.cad_elements)
    total_groups = sum(1 for g in vm.CadGroups if not g.is_beam())
    vm.Status = "Them {} elements moi → tong {} elements, {} nhom.".format(
        len(elements), total_elems, total_groups
    )

    _redraw(window)
    MessageBox.Show(
        "Da them {} elements tu AutoCAD.\nTong cong: {} elements ({} polyline), {} nhom.".format(
            len(raw_elements), total_elems, len(elements), total_groups
        ),
        "Hoan thanh"
    )


def on_select_grid_in_cad(sender, e):
    """Chọn 1 đường tham chiếu trong CAD (StartPoint = gốc, phương chiều = hướng transform)."""
    window = sender.Tag
    vm = window.DataContext

    doc = get_acad_doc()
    if not doc:
        MessageBox.Show("Khong the ket noi AutoCAD.", "Loi ket noi")
        return

    grids = select_grid_in_cad(doc)
    if len(grids) < 1:
        MessageBox.Show(
            "Can chon dung 1 duong tham chieu trong CAD.",
            "Chon tham chieu CAD"
        )
        return

    vm.cad_grid_elements = grids[:1]
    vm.Status = "Da chon duong tham chieu CAD. StartPoint: ({:.1f}, {:.1f}) mm".format(
        grids[0].points[0][0], grids[0].points[0][1]
    )

    _redraw(window)
    MessageBox.Show("Da chon duong tham chieu trong AutoCAD.", "Hoan thanh")


def on_select_grid_in_revit(sender, e):
    """Chọn 2 Grid trong Revit – raise ExternalEvent."""
    window = sender.Tag
    if hasattr(window, '_ext_select_grid'):
        window._ext_select_grid.Raise()
    else:
        MessageBox.Show(
            "ExternalEvent chua duoc khoi tao. Hay khoi dong lai tool.",
            "Loi"
        )


def on_create_model(sender, e):
    """Tạo model trong Revit – raise ExternalEvent."""
    window = sender.Tag
    vm = window.DataContext

    if not vm.cad_elements and not vm.beam_elements:
        MessageBox.Show("Chua chon elements tu CAD (column/footing hoac beam).", "Thieu du lieu")
        return

    if hasattr(window, '_ext_create_model'):
        window._ext_create_model.Raise()
    else:
        MessageBox.Show(
            "ExternalEvent chua duoc khoi tao. Hay khoi dong lai tool.",
            "Loi"
        )


# ============================================================
#   CANVAS RESIZE + REDRAW
# ============================================================
def _redraw(window):
    """Vẽ lại canvas theo kích thước hiện tại và selection state."""
    canvas = window.DrawingCanvas
    w = canvas.ActualWidth
    h = canvas.ActualHeight
    if w < 10 or h < 10:
        return

    vm = window.DataContext

    # Individual element selection (canvas click) – ưu tiên hơn group selection
    selected_elem  = getattr(vm, 'SelectedElement', None)
    selected_label = None
    if selected_elem is None:
        # Group selection (DataGrid)
        sel = getattr(vm, 'SelectedGroup', None)
        if sel is not None:
            selected_label = getattr(sel, 'label', None)

    all_elements = list(getattr(vm, 'cad_elements', [])) + list(getattr(vm, 'beam_elements', []))
    _draw_elements_on_canvas(
        canvas,
        all_elements,
        getattr(vm, 'cad_grid_elements', []),
        w, h,
        selected_label=selected_label,
        selected_elem=selected_elem,
    )


def on_canvas_size_changed(sender, e):
    try:
        window = sender.Tag
        _redraw(window)
    except Exception:
        pass
