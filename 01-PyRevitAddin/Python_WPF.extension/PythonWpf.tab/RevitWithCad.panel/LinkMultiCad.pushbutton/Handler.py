# -*- coding: utf-8 -*-
"""
Handler.py - Xử lý sự kiện cho cửa sổ Link Multiple CAD.

Chức năng chính:
  - SelectBaseLineHandler : IExternalEventHandler – chọn đường tham chiếu trong Revit
  - LinkCadHandler        : IExternalEventHandler – link file DWG + tính transform
  - bind_handlers(window) : gắn tất cả button/event cho WPF window
"""
import math
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')

from clr import Reference
from System.Windows.Forms import (
    FolderBrowserDialog, DialogResult as WFDialogResult,
    MessageBox, MessageBoxButtons, MessageBoxIcon
)

# ============================================================
#   REVIT API
# ============================================================
try:
    from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
    from Autodesk.Revit.UI.Selection import ObjectType
    from Autodesk.Revit.DB import (
        FilteredElementCollector, Level, Transaction,
        XYZ, Line as RvtLine, ElementId,
        ImportInstance, DWGImportOptions, ImportPlacement,
        Options, GeometryInstance, BuiltInParameter,
        ElementTransformUtils
    )
    # DWGColorMode có thể không tồn tại trong một số phiên bản Revit
    try:
        from Autodesk.Revit.DB import DWGColorMode
        _HAS_COLOR_MODE = True
    except Exception:
        _HAS_COLOR_MODE = False

    _REVIT_AVAILABLE = True
except Exception:
    _REVIT_AVAILABLE = False


# ============================================================
#   GEOMETRY HELPERS
# ============================================================

def _get_curve_from_element(elem):
    """
    Lấy Curve từ một Revit element (Grid, ModelLine, DetailLine, ...).
    Trả về Curve hoặc raise Exception nếu không tìm thấy.
    """
    # Thử Location.Curve (ModelLine, DetailLine, CurveElement...)
    try:
        loc = elem.Location
        if loc is not None:
            curve = loc.Curve
            if curve is not None:
                return curve
    except Exception:
        pass

    # Thử thuộc tính .Curve trực tiếp (Grid)
    try:
        curve = elem.Curve
        if curve is not None:
            return curve
    except Exception:
        pass

    raise Exception(
        "Không lấy được Curve từ element Id={}.".format(elem.Id)
    )


def _find_origin_line(doc, import_inst, layer_name):
    """
    Tìm đường thẳng đầu tiên trên layer *layer_name* trong ImportInstance.
    Duyệt đệ quy qua GeometryInstance.
    Trả về Revit Line hoặc None nếu không tìm thấy.
    """
    geom_opts = Options()
    geom_opts.ComputeReferences = False
    geom_element = import_inst.get_Geometry(geom_opts)
    if geom_element is None:
        return None

    target_layer = layer_name.strip().lower()

    def _search(geom_iter):
        for geom_obj in geom_iter:
            # Nếu là GeometryInstance (block / group) → đệ quy vào bên trong
            if isinstance(geom_obj, GeometryInstance):
                result = _search(geom_obj.GetInstanceGeometry())
                if result is not None:
                    return result
                continue

            # Kiểm tra layer của đối tượng hình học
            style_id = geom_obj.GraphicsStyleId
            if style_id == ElementId.InvalidElementId:
                continue
            style = doc.GetElement(style_id)
            if style is None:
                continue
            try:
                cat = style.GraphicsStyleCategory
                if cat is None:
                    continue
                if cat.Name.strip().lower() != target_layer:
                    continue
            except Exception:
                continue

            # Chỉ lấy Line (đường thẳng), không lấy Arc / PolyLine
            if isinstance(geom_obj, RvtLine):
                return geom_obj

        return None

    return _search(geom_element)


def _apply_transform(doc, import_id, cad_line, revit_baseline):
    """
    Xoay và tịnh tiến ImportInstance để đường *cad_line* khớp với *revit_baseline*.

    Bước 1 – Xoay:
        Xoay quanh trục thẳng đứng đi qua điểm đầu (start) của cad_line
        một góc theta = angle(Revit baseline) − angle(cad_line).
        → Điểm đầu cad_line không dịch chuyển; hướng khớp với Revit baseline.

    Bước 2 – Tịnh tiến:
        Dịch toàn bộ ImportInstance sao cho điểm đầu cad_line
        trùng với điểm đầu Revit baseline.
    """
    # ---- Lấy thông tin CAD line ----
    cad_start = cad_line.GetEndPoint(0)
    cad_dir   = cad_line.Direction       # unit vector

    # ---- Lấy thông tin Revit baseline ----
    revit_curve = _get_curve_from_element(revit_baseline)
    revit_start = revit_curve.GetEndPoint(0)
    revit_dir   = revit_curve.Direction  # unit vector

    # ---- Tính góc xoay ----
    theta = (
        math.atan2(revit_dir.Y, revit_dir.X)
        - math.atan2(cad_dir.Y, cad_dir.X)
    )

    # ---- Bước 1: Xoay quanh trục đứng qua cad_start ----
    rot_axis = RvtLine.CreateBound(
        cad_start,
        XYZ(cad_start.X, cad_start.Y, cad_start.Z + 1.0)
    )
    ElementTransformUtils.RotateElement(doc, import_id, rot_axis, theta)

    # Sau khi xoay quanh cad_start, điểm đầu của đường CAD vẫn là cad_start.
    # Bước 2: Tịnh tiến sao cho cad_start → revit_start
    delta = revit_start - cad_start
    ElementTransformUtils.MoveElement(doc, import_id, delta)


def _link_dwg(doc, file_path, options, view):
    """
    Link (hoặc import nếu Link không khả dụng) file DWG vào Revit.
    Sử dụng phương thức out-parameter thông qua clr.Reference.
    Ném Exception nếu cả hai cách đều thất bại.
    """
    # Thử Link trước (giữ lại external reference)
    try:
        link_id_ref = Reference[ElementId](ElementId.InvalidElementId)
        doc.Link(file_path, options, view, link_id_ref)
        return
    except Exception:
        pass

    # Fallback: Import (embedded)
    try:
        imp_id_ref = Reference[ElementId](ElementId.InvalidElementId)
        doc.Import(file_path, options, view, imp_id_ref)
        return
    except Exception as ex:
        raise Exception("Không thể Link/Import '{}': {}".format(file_path, str(ex)))


# ============================================================
#   REVIT EXTERNAL EVENT HANDLERS
# ============================================================

class SelectBaseLineHandler(IExternalEventHandler if _REVIT_AVAILABLE else object):
    """
    Cho người dùng click chọn 1 đường trong Revit
    (Grid, ModelLine, DetailLine, ...) làm đường tham chiếu baseline.
    Kết quả được lưu vào vm.revit_baseline.
    """

    def __init__(self, vm):
        self.vm = vm

    def Execute(self, uiapp):
        uiDoc = uiapp.ActiveUIDocument
        vm    = self.vm
        try:
            ref  = uiDoc.Selection.PickObject(
                ObjectType.Element,
                "Chọn đường tham chiếu (Grid / Line) trong Revit..."
            )
            elem = uiDoc.Document.GetElement(ref)

            # Validate: phải lấy được Curve
            _get_curve_from_element(elem)

            vm.revit_baseline = elem
            vm.Status = u"Baseline đã chọn: Id = {}".format(elem.Id)

        except Exception as ex:
            msg = str(ex)
            if 'cancel' in msg.lower() or 'OperationCanceledException' in msg:
                vm.Status = u"Đã hủy chọn baseline."
            else:
                vm.Status = u"Lỗi chọn baseline: {}".format(msg)

    def GetName(self):
        return "SelectBaseLineHandler"


class LinkCadHandler(IExternalEventHandler if _REVIT_AVAILABLE else object):
    """
    Với mỗi CadFileRow đã có Base Level:
      1. Link DWG vào Revit ở Internal Origin.
      2. Đọc geometry, tìm đường trên layer LayerOriginName.
      3. Tính transform (rotate + translate) so với Revit baseline.
      4. Áp dụng transform lên ImportInstance.
      5. Gán Base Level.
    """

    def __init__(self, vm):
        self.vm = vm

    def Execute(self, uiapp):
        uiDoc = uiapp.ActiveUIDocument
        doc   = uiDoc.Document
        vm    = self.vm

        # ---- Validate ----
        layer_name = (vm.LayerOriginName or '').strip()
        if not layer_name:
            vm.Status = u"Lỗi: Chưa nhập Layer Origin Name."
            return

        if not vm.has_baseline():
            vm.Status = u"Lỗi: Chưa chọn Baseline trong Revit."
            return

        rows = [r for r in vm.CadFiles if r.is_ready()]
        if not rows:
            vm.Status = u"Lỗi: Không có file nào đã chọn Base Level."
            return

        # ---- Resolve Revit Levels ----
        levels_map = {}
        try:
            for lv in FilteredElementCollector(doc).OfClass(Level).ToElements():
                try:
                    levels_map[lv.Name] = lv
                except Exception:
                    pass
        except Exception:
            pass

        # ---- DWG options ----
        options = DWGImportOptions()
        options.Placement = ImportPlacement.Origin
        if _HAS_COLOR_MODE:
            try:
                options.ColorMode = DWGColorMode.BlackAndWhite
            except Exception:
                pass

        active_view = uiDoc.ActiveView
        success_list = []
        error_list   = []

        for row in rows:
            try:
                with Transaction(doc, u'Link CAD: {}'.format(row.FileName)) as t:
                    t.Start()

                    # Ghi nhớ các ImportInstance hiện có
                    before_ids = set(
                        FilteredElementCollector(doc)
                        .OfClass(ImportInstance)
                        .ToElementIds()
                    )

                    # Link DWG
                    _link_dwg(doc, row.FilePath, options, active_view)

                    # Tìm element mới
                    after_ids = set(
                        FilteredElementCollector(doc)
                        .OfClass(ImportInstance)
                        .ToElementIds()
                    )
                    new_ids = after_ids - before_ids

                    if not new_ids:
                        t.RollBack()
                        error_list.append(
                            u"{}: Không tạo được ImportInstance.".format(row.FileName)
                        )
                        continue

                    import_id   = list(new_ids)[0]
                    import_inst = doc.GetElement(import_id)

                    # Unpin để phép rotate + move
                    try:
                        if import_inst.Pinned:
                            import_inst.Pinned = False
                    except Exception:
                        pass

                    # Tìm đường origin theo layer
                    origin_line = _find_origin_line(doc, import_inst, layer_name)
                    if origin_line is None:
                        t.RollBack()
                        error_list.append(
                            u"{}: Không tìm thấy layer '{}'.".format(
                                row.FileName, layer_name
                            )
                        )
                        continue

                    # Áp dụng transform
                    _apply_transform(doc, import_id, origin_line, vm.revit_baseline)

                    # Gán Base Level
                    lv = levels_map.get(row.BaseLevel)
                    if lv is not None:
                        try:
                            p = import_inst.get_Parameter(
                                BuiltInParameter.IMPORT_BASE_LEVEL
                            )
                            if p is not None and not p.IsReadOnly:
                                p.Set(lv.Id)
                        except Exception:
                            pass

                    t.Commit()
                    success_list.append(row.FileName)

            except Exception as ex:
                error_list.append(u"{}: {}".format(row.FileName, str(ex)))

        # ---- Cập nhật status ----
        parts = []
        if success_list:
            parts.append(u"Đã link {} file.".format(len(success_list)))
        if error_list:
            parts.append(u"Lỗi [{}]: {}".format(
                len(error_list), u" | ".join(error_list)
            ))
        vm.Status = u"  ".join(parts) if parts else u"Hoàn thành."

    def GetName(self):
        return "LinkCadHandler"


# ============================================================
#   BUTTON EVENT HANDLERS
# ============================================================

def on_select_folder(sender, e):
    """
    Mở FolderBrowserDialog, quét .dwg và cập nhật AvailableDwgFiles
    (được dùng bởi ComboBox cột File Cad trong DataGrid).
    """
    window = _get_window(sender)
    if window is None:
        return
    vm = window.DataContext

    dlg = FolderBrowserDialog()
    dlg.Description = "Chọn thư mục chứa các file CAD (.dwg)"
    if vm.folder_path:
        dlg.SelectedPath = vm.folder_path

    if dlg.ShowDialog() == WFDialogResult.OK:
        count = vm.scan_folder(dlg.SelectedPath)
        vm.Status = u"Folder: {} | Tìm thấy {} file .dwg.".format(
            dlg.SelectedPath, count
        )
    else:
        vm.Status = u"Đã hủy chọn folder."


def on_add(sender, e):
    """
    Thêm 1 hàng trống vào DataGrid.
    Nếu chưa chọn folder (AvailableDwgFiles rỗng), tự động mở dialog chọn folder trước.
    """
    window = _get_window(sender)
    if window is None:
        return
    vm = window.DataContext

    # Nếu danh sách file chưa có, nhắc chọn folder trước
    if len(vm.AvailableDwgFiles) == 0:
        on_select_folder(sender, e)
        if len(vm.AvailableDwgFiles) == 0:
            vm.Status = u"Chưa có file .dwg nào. Hãy nhấn 'Cad path folder' trước."
            return

    vm.add_empty_row()
    vm.Status = u"Đã thêm 1 dòng. Chọn file .dwg và Base Level cho dòng mới."


def on_remove(sender, e):
    """
    Xóa tất cả các dòng đang được chọn trong DataGrid (hỗ trợ multi-select bằng Ctrl).
    """
    window = _get_window(sender)
    if window is None:
        return
    vm = window.DataContext

    # Đọc SelectedItems từ DataGrid (hỗ trợ Extended selection)
    try:
        selected = list(window.DgFiles.SelectedItems)
    except Exception:
        selected = []

    if not selected:
        vm.Status = u"Chưa chọn dòng nào để xóa."
        return

    names = []
    for r in selected:
        try:
            names.append(r.FileName)
        except Exception:
            pass

    vm.remove_rows(selected)
    vm.Status = u"Đã xóa {} dòng: {}".format(
        len(names), u', '.join(names)
    )


def on_select_baseline(sender, e):
    """Kích hoạt ExternalEvent để chọn đường tham chiếu trong Revit."""
    window = _get_window(sender)
    if window is None:
        return
    try:
        window._ext_select_baseline.Raise()
    except Exception as ex:
        window.DataContext.Status = u"Lỗi ExternalEvent: {}".format(str(ex))


def on_link(sender, e):
    """Kích hoạt ExternalEvent để link tất cả file DWG."""
    window = _get_window(sender)
    if window is None:
        return
    vm = window.DataContext

    if not vm.has_baseline():
        vm.Status = u"Chưa chọn Baseline trong Revit. Hãy nhấn 'Select Base Line in Revit' trước."
        return

    if not (vm.LayerOriginName or '').strip():
        vm.Status = u"Chưa nhập Layer Origin Name."
        return

    rows_ready = [r for r in vm.CadFiles if r.is_ready()]
    if not rows_ready:
        vm.Status = u"Không có file nào đã chọn Base Level."
        return

    vm.Status = u"Đang link {} file...".format(len(rows_ready))
    try:
        window._ext_link_cad.Raise()
    except Exception as ex:
        vm.Status = u"Lỗi ExternalEvent: {}".format(str(ex))


# ============================================================
#   BIND HANDLERS
# ============================================================

def bind_handlers(window):
    """
    Gắn tất cả button Click events cho WPF window.
    Được gọi 1 lần trong ModelByCadWindow.__init__.
    """
    window.BtnCadFolder.Click      += on_select_folder
    window.BtnAdd.Click            += on_add
    window.BtnRemove.Click         += on_remove
    window.BtnSelectBaseLine.Click += on_select_baseline
    window.BtnLink.Click           += on_link
    # Wire DataGrid LoadingRow để gắn SelectionChanged cho ComboBox trong mỗi dòng
    window.DgFiles.LoadingRow      += _on_dg_loading_row


# ============================================================
#   DATAGRID ROW – wire ComboBox SelectionChanged qua LoadingRow
# ============================================================

def _on_dg_loading_row(sender, e):
    """
    Fires mỗi khi DataGrid chuẩn bị 1 row.
    Đăng ký Loaded trên row để chờ visual tree sẵn sàng rồi mới wire ComboBox.
    """
    e.Row.Loaded += _on_dg_row_loaded


def _on_dg_row_loaded(sender, e):
    """
    Fires sau khi DataGridRow đã render xong.
    Tìm tất cả ComboBox bên trong row và gắn SelectionChanged.
    Tag 'filecad' → cập nhật SelectedEntry; Tag 'level' → cập nhật BaseLevel.
    Dùng 'wired' trong Tag để tránh đăng ký trùng khi row bị tái dùng (virtualization).
    """
    from System.Windows.Controls import ComboBox
    combos = _find_visual_children(sender, ComboBox)
    for combo in combos:
        tag = combo.Tag
        # Tag là string 'filecad' hoặc 'level' – chưa wire
        if tag in ('filecad', 'level'):
            combo.SelectionChanged += _on_combo_selection_changed
            # Đánh dấu đã wire để tránh đăng ký lại
            combo.Tag = tag + '_wired'


def _on_combo_selection_changed(sender, e):
    """
    Cập nhật trực tiếp property của CadFileRow khi người dùng chọn item.
    DataContext của ComboBox là CadFileRow tương ứng.
    """
    combo = sender
    item  = combo.SelectedItem
    if item is None:
        return
    try:
        row = combo.DataContext
        if row is None:
            return
        tag = combo.Tag or ''
        if 'filecad' in tag:
            row.SelectedEntry = item
        elif 'level' in tag:
            row.BaseLevel = item
    except Exception:
        pass


def _find_visual_children(parent, child_type):
    """
    Duyệt đệ quy visual tree, trả về list các node có kiểu *child_type*.
    """
    from System.Windows.Media import VisualTreeHelper
    results = []
    try:
        count = VisualTreeHelper.GetChildrenCount(parent)
    except Exception:
        return results
    for i in range(count):
        try:
            child = VisualTreeHelper.GetChild(parent, i)
        except Exception:
            continue
        if isinstance(child, child_type):
            results.append(child)
        results.extend(_find_visual_children(child, child_type))
    return results


# ============================================================
#   PRIVATE HELPERS
# ============================================================

def _get_window(sender):
    """
    Lấy WPF Window từ sender (button hoặc bất kỳ DependencyObject nào).
    Dùng Window.GetWindow() – cách chính thống của WPF.
    """
    try:
        from System.Windows import Window as WpfWindow
        return WpfWindow.GetWindow(sender)
    except Exception:
        pass
    return None
