# -*- coding: utf-8 -*-
"""
CadUtils.py – Tiện ích đọc/phân tích geometry từ AutoCAD COM.

Exports chính:
    CadElement              – dataclass geometry 1 element CAD
    _MergedPolyline         – pseudo-element từ merge nhiều segment
    CadBeamPair             – 1 cặp line + text → dầm
    get_acad_doc()          – kết nối AutoCAD ActiveDocument
    load_file_to_doc(path)  – mở 1 DWG file, trả về Document COM
    extract_all_from_doc(doc) – trích xuất toàn bộ elements + layers
    filter_elements_by_rules(elements, rules) – lọc theo OR-rule list
    select_grid_in_cad(doc) – chọn 1 đường tham chiếu (interactive)
    merge_lines_to_closed_polylines(elements) – gom line thành closed poly
    group_elements_by_label(elements)         – nhóm theo kích thước
    select_beam_elements_in_cad(doc)          – chọn line+text cho dầm
    group_beam_pairs_by_label(pairs)          – nhóm dầm theo kích thước
    analyze_condition(elements, category)     – phân tích theo category
"""
import math
import re as _re
import clr
from System.Runtime.InteropServices import Marshal
from System import Type, Activator


# =============================================
# DATA CLASSES
# =============================================
class CadElement:
    """
    Lưu geometry đã parse của 1 đối tượng CAD.
    type   : 'polyline' | 'line' | 'circle' | 'arc' | 'unknown'
    points : list of (x, y)
    center : (x, y)   – circle / arc
    radius : float    – circle / arc
    layer  : str      – layer name (uppercase)
    obj    : COM ref gốc
    """
    def __init__(self, obj):
        self.obj         = obj
        self.points      = []
        self.center      = None
        self.radius      = 0.0
        self.start_angle = 0.0
        self.end_angle   = 0.0
        self.type        = 'unknown'
        self.layer       = ''
        self._parse(obj)

    def _parse(self, obj):
        try:
            name = obj.ObjectName.upper()

            # Layer
            try:
                self.layer = (obj.Layer or '').upper()
            except Exception:
                pass

            if 'CIRCLE' in name:
                self.type = 'circle'
                c = obj.Center
                self.center = (c[0], c[1])
                self.radius = obj.Radius

            elif 'ARC' in name and 'POLYLINE' not in name:
                self.type = 'arc'
                c = obj.Center
                self.center = (c[0], c[1])
                self.radius = obj.Radius
                self.start_angle = obj.StartAngle
                self.end_angle   = obj.EndAngle

            elif 'LINE' in name:
                if 'POLYLINE' not in name:
                    sp = obj.StartPoint
                    ep = obj.EndPoint
                    self.points = [(sp[0], sp[1]), (ep[0], ep[1])]
                    self.type = 'line'
                else:
                    coords = obj.Coordinates
                    step = 2 if len(coords) % 2 == 0 else 3
                    n = len(coords) // step
                    self.points = [(coords[i*step], coords[i*step+1]) for i in range(n)]
                    self.type = 'polyline'
                    try:
                        if obj.Closed and len(self.points) >= 2:
                            p0, pn = self.points[0], self.points[-1]
                            if (p0[0]-pn[0])**2 + (p0[1]-pn[1])**2 > 1e-6:
                                self.points.append(p0)
                    except Exception:
                        pass
            else:
                self.type = 'unknown'
        except Exception as ex:
            print('CadElement parse error: {}'.format(ex))
            self.type = 'unknown'


class _MergedPolyline:
    """Pseudo-CadElement: polyline ghép từ nhiều line/segment."""
    def __init__(self, points, layer=''):
        self.obj         = None
        self.type        = 'polyline'
        self.points      = points
        self.center      = None
        self.radius      = 0.0
        self.start_angle = 0.0
        self.end_angle   = 0.0
        self.layer       = layer


class CadBeamPair:
    """
    1 dầm CAD = 1 location line + text kích thước gần nhất.
    Attributes: start, end, width, height, label, raw_text, line_elem, layer
    """
    def __init__(self, line_elem, text_str, width, height):
        self.line_elem = line_elem
        self.start     = line_elem.points[0]
        self.end       = line_elem.points[-1]
        self.raw_text  = text_str
        self.width     = width
        self.height    = height
        self.label     = 'BEA: {}x{}'.format(width, height)
        self.type      = 'beam_line'
        self.points    = line_elem.points
        self.layer     = getattr(line_elem, 'layer', '')


# =============================================
# AUTO-CAD CONNECTION
# =============================================
def get_acad_doc():
    """Kết nối tới AutoCAD đang mở, trả về ActiveDocument hoặc None."""
    try:
        acad = Marshal.GetActiveObject("AutoCAD.Application")
        return acad.ActiveDocument
    except Exception:
        try:
            acad_type = Type.GetTypeFromProgID("AutoCAD.Application")
            acad      = Activator.CreateInstance(acad_type)
            return acad.ActiveDocument
        except Exception as ex:
            print("Khong the ket noi AutoCAD: {}".format(ex))
            return None


def load_file_to_doc(filepath):
    """
    Mở file DWG qua AutoCAD COM.
    Trả về Document COM hoặc None nếu thất bại.
    """
    try:
        acad = Marshal.GetActiveObject("AutoCAD.Application")
    except Exception:
        try:
            acad_type = Type.GetTypeFromProgID("AutoCAD.Application")
            acad      = Activator.CreateInstance(acad_type)
        except Exception as ex:
            print("load_file_to_doc: Khong the ket noi AutoCAD: {}".format(ex))
            return None
    try:
        doc = acad.Documents.Open(filepath, False)   # False = không read-only
        return doc
    except Exception as ex:
        print("load_file_to_doc: Khong the mo file '{}': {}".format(filepath, ex))
        return None


# =============================================
# EXTRACT ALL FROM DOCUMENT
# =============================================
def extract_all_from_doc(doc):
    """
    Trích xuất toàn bộ elements + layers + texts từ Document COM.
    Trả về (elements: list[CadElement], layers: list[str],
             texts: list[(content, x, y, layer)]).
    """
    elements = []
    layers   = set()
    texts    = []   # (content, x, y, layer)

    try:
        model_space = doc.ModelSpace
        count = model_space.Count
        for i in range(count):
            try:
                obj  = model_space.Item(i)
                name = ''
                try: name = obj.ObjectName.upper()
                except Exception: pass

                # Parse TEXT / MTEXT trước khi thử CadElement
                if 'TEXT' in name:
                    try:
                        content = obj.TextString
                        ins     = obj.InsertionPoint
                        lyr     = ''
                        try: lyr = (obj.Layer or '').upper()
                        except Exception: pass
                        texts.append((content, ins[0], ins[1], lyr))
                        layers.add(lyr)
                    except Exception:
                        pass
                    continue

                elem = CadElement(obj)
                if elem.type != 'unknown':
                    elements.append(elem)
                if elem.layer:
                    layers.add(elem.layer)
            except Exception:
                pass

        # Thêm layer từ Layers collection
        try:
            for j in range(doc.Layers.Count):
                try:
                    layers.add(doc.Layers.Item(j).Name.upper())
                except Exception:
                    pass
        except Exception:
            pass

    except Exception as ex:
        print("extract_all_from_doc error: {}".format(ex))

    return elements, sorted(layers), texts


# =============================================
# INTERACTIVE SELECTION (vẫn dùng cho grid ref)
# =============================================
def _raw_select(doc, prompt_msg):
    raw = []
    ssa_name = "PySel_ModelByCad"
    try:
        doc.Utility.Prompt(prompt_msg)
        try:
            doc.SelectionSets.Item(ssa_name).Delete()
        except Exception:
            pass
        ssa = doc.SelectionSets.Add(ssa_name)
        ssa.SelectOnScreen()
        for i in range(ssa.Count):
            raw.append(ssa.Item(i))
        ssa.Delete()
    except Exception as ex:
        print("Loi chon doi tuong: {}".format(ex))
    return raw


def select_elements_in_cad(doc):
    """Interactive: chọn elements trên màn hình CAD."""
    raw = _raw_select(doc, "\nChon cac elements va nhan Enter...\n")
    result = []
    for obj in raw:
        try:
            elem = CadElement(obj)
            if elem.type != 'unknown':
                result.append(elem)
        except Exception as ex:
            print("Bo qua: {}".format(ex))
    return result


def select_grid_in_cad(doc):
    """Interactive: chọn 1 đường tham chiếu trong CAD."""
    raw    = _raw_select(doc, "\nChon DUNG 1 duong tham chieu trong CAD roi nhan Enter...\n")
    result = []
    for obj in raw:
        try:
            elem = CadElement(obj)
            if elem.type in ('line', 'polyline') and len(result) < 1:
                result.append(elem)
        except Exception as ex:
            print("Bo qua: {}".format(ex))
    if len(result) != 1:
        print("Canh bao: can chon dung 1 duong, hien co {}.".format(len(result)))
    return result


# =============================================
# FILTER BY RULES  (OR logic)
# =============================================
def _poly_length(elem):
    """Tổng độ dài các segment của polyline / line (mm)."""
    pts = getattr(elem, 'points', [])
    if len(pts) < 2:
        return 0.0
    total = 0.0
    for i in range(len(pts) - 1):
        dx = pts[i+1][0] - pts[i][0]
        dy = pts[i+1][1] - pts[i][1]
        total += math.sqrt(dx*dx + dy*dy)
    return total


def _poly_width(elem):
    """Chiều rộng hướng ngang của bounding box (dùng cho Distance rule)."""
    pts = getattr(elem, 'points', [])
    if not pts:
        return 0.0
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    w  = max(xs) - min(xs)
    h  = max(ys) - min(ys)
    return min(w, h)   # chiều nhỏ hơn ~ khoảng cách 2 line song song


def _apply_rule(elem, rule_dict):
    """
    Kiểm tra 1 element có thỏa 1 rule không.
    rule_dict: {'parameter': str, 'ruler': str, 'value': str}

    Các parameter được hỗ trợ:
      Layer Name           → so sánh layer name (Equal)
      Length               → chiều dài element (is greater than / is less than)
      Min Beam Distance    → dùng để trích min_d khi gọi detect_beams; không filter element
      Max Beam Distance    → dùng để trích max_d khi gọi detect_beams; không filter element
    """
    param = rule_dict.get('parameter', '')
    ruler = rule_dict.get('ruler', '')
    value = rule_dict.get('value', '')

    if param == 'Layer Name':
        return elem.layer.upper() == value.upper()

    # Các parameter điều khiển thuật toán, không dùng để filter element
    if param in ('Min Beam Distance', 'Max Beam Distance', 'Text Layer'):
        return False

    try:
        num_value = float(value)
    except (ValueError, TypeError):
        return False

    if param == 'Length':
        length = _poly_length(elem)
        if ruler == 'is greater than':
            return length > num_value
        if ruler == 'is less than':
            return length < num_value
        if ruler == 'Equal':
            return abs(length - num_value) < 1.0

    return False


def filter_elements_by_rules(elements, rules):
    """
    Lọc elements theo list rules với logic OR:
    element thỏa BẤT KỲ 1 rule nào → được chọn.

    rules: list[RuleRow] hoặc list[dict]
    Trả về list[element].
    """
    if not rules:
        return list(elements)

    # Chuẩn hóa sang dict
    rule_dicts = []
    for r in rules:
        if hasattr(r, 'to_dict'):
            rule_dicts.append(r.to_dict())
        elif isinstance(r, dict):
            rule_dicts.append(r)

    if not rule_dicts:
        return list(elements)

    result = []
    for elem in elements:
        for rd in rule_dicts:
            if _apply_rule(elem, rd):
                result.append(elem)
                break
    return result


# =============================================
# MERGE LINES → CLOSED POLYLINES
# =============================================
def merge_lines_to_closed_polylines(elements, tolerance=1.0):
    """
    Gom các line/polyline có endpoints chung thành polyline kín.
    Các element khác (circle, arc) được giữ nguyên.
    """
    tol_sq = tolerance ** 2

    def _near(a, b):
        return (a[0]-b[0])**2 + (a[1]-b[1])**2 <= tol_sq

    line_elems  = [e for e in elements if e.type in ('line', 'polyline')]
    other_elems = [e for e in elements if e.type not in ('line', 'polyline')]

    if not line_elems:
        return list(elements)

    already_closed = []
    segments       = []
    seg_layers     = []   # layer tương ứng với mỗi segment

    for elem in line_elems:
        pts = elem.points
        if len(pts) < 2:
            continue
        if len(pts) >= 3 and _near(pts[0], pts[-1]):
            mp = _MergedPolyline(list(pts), elem.layer)
            already_closed.append(mp)
        else:
            for i in range(len(pts) - 1):
                segments.append((pts[i], pts[i+1]))
                seg_layers.append(elem.layer)

    if not segments:
        return already_closed + other_elems

    used = [False] * len(segments)

    def _find_next(pt):
        for i, (sa, sb) in enumerate(segments):
            if used[i]:
                continue
            if _near(pt, sa):
                return i, sb
            if _near(pt, sb):
                return i, sa
        return None, None

    chains = []
    chain_layers = []

    for start_idx in range(len(segments)):
        if used[start_idx]:
            continue
        sa0, sb0 = segments[start_idx]
        chain = [sa0, sb0]
        clyr  = seg_layers[start_idx]
        used[start_idx] = True

        while True:
            idx, nxt = _find_next(chain[-1])
            if idx is None:
                break
            used[idx] = True
            chain.append(nxt)
            if _near(chain[0], chain[-1]):
                chain[-1] = chain[0]
                break
            if len(chain) > len(segments) + 2:
                break

        if not _near(chain[0], chain[-1]):
            while True:
                idx, nxt = _find_next(chain[0])
                if idx is None:
                    break
                used[idx] = True
                chain.insert(0, nxt)
                if _near(chain[0], chain[-1]):
                    chain[-1] = chain[0]
                    break
                if len(chain) > len(segments) + 2:
                    break

        chains.append(chain)
        chain_layers.append(clyr)

    merged = already_closed + [
        _MergedPolyline(ch, chain_layers[i]) for i, ch in enumerate(chains)
    ]
    return merged + other_elems


# =============================================
# GEOMETRY ANALYSIS HELPERS
# =============================================
def _round5(value):
    return int(round(value / 5.0) * 5)


def _mm(v):
    return v  # CAD units = mm


def _dist_sq(p1, p2):
    return (p1[0]-p2[0])**2 + (p1[1]-p2[1])**2


def _midpoint(p1, p2):
    return ((p1[0]+p2[0])/2.0, (p1[1]+p2[1])/2.0)


def _bbox(pts):
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def _poly_length(elem):
    pts = getattr(elem, 'points', [])
    if len(pts) < 2:
        return 0.0
    total = 0.0
    for i in range(len(pts)-1):
        dx = pts[i+1][0] - pts[i][0]
        dy = pts[i+1][1] - pts[i][1]
        total += math.sqrt(dx*dx + dy*dy)
    return total


def _is_closed(elem, tol=1.0):
    """Kiểm tra polyline có kín không."""
    if elem.type == 'circle':
        return True
    pts = elem.points
    if len(pts) < 3:
        return False
    return _dist_sq(pts[0], pts[-1]) <= tol**2


def _aspect_ratio(pts):
    """Tính aspect ratio của bounding box (max/min)."""
    mn_x, mn_y, mx_x, mx_y = _bbox(pts)
    w = mx_x - mn_x or 1e-6
    h = mx_y - mn_y or 1e-6
    return max(w, h) / min(w, h)


# =============================================
# ANALYZE PER CATEGORY
# =============================================
def analyze_condition(elements, category):
    """
    Phân tích elements theo Category, trả về list elements phù hợp.
    category: str (Structural Columns / Structural Framing / Walls / Structural Foundations)
    """
    cat_lower = category.lower().strip()

    if 'column' in cat_lower:
        return _analyze_columns(elements)
    if 'framing' in cat_lower or 'beam' in cat_lower:
        return detect_beams_from_lines(elements)
    if 'wall' in cat_lower:
        return _analyze_walls(elements)
    if 'foundation' in cat_lower or 'footing' in cat_lower:
        return _analyze_columns(elements)  # móng xử lý tương tự cột

    # Fallback: trả về tất cả
    return list(elements)


def _analyze_columns(elements):
    """
    Phân tích cột/móng:
    - Closed polyline + circle
    - Aspect ratio < 3
    - Loại hatch/block trang trí (area quá nhỏ)
    """
    result = []
    for elem in elements:
        if elem.type == 'circle':
            if elem.radius > 50:  # loại circle rất nhỏ (hatch)
                result.append(elem)
            continue

        if elem.type in ('polyline',) and _is_closed(elem):
            pts = elem.points
            if len(pts) < 3:
                continue
            ratio = _aspect_ratio(pts)
            if ratio >= 3:
                continue  # không phải cột, có thể là dầm kín
            mn_x, mn_y, mx_x, mx_y = _bbox(pts)
            w = mx_x - mn_x
            h = mx_y - mn_y
            if w < 50 or h < 50:
                continue  # quá nhỏ → hatch/decoration
            result.append(elem)

    return result


def _analyze_framings(elements):
    """
    Phân tích dầm:
    - Open line/polyline (không kín)
    - Loại closed polyline, circle
    - Giữ lại line đủ dài (> 100mm)
    """
    result = []
    for elem in elements:
        if elem.type == 'circle':
            continue
        if elem.type == 'polyline' and _is_closed(elem):
            continue
        if elem.type in ('line', 'polyline'):
            length = _poly_length(elem)
            if length < 100:
                continue
            result.append(elem)
    return result


def _analyze_walls(elements):
    """
    Phân tích tường:
    - Open line/polyline
    - Closed polyline có aspect ratio > 3
    - Loại segment quá ngắn
    """
    result = []
    for elem in elements:
        if elem.type == 'circle':
            continue
        if elem.type == 'polyline' and _is_closed(elem):
            pts = elem.points
            ratio = _aspect_ratio(pts)
            if ratio > 3:
                result.append(elem)
            continue
        if elem.type in ('line', 'polyline'):
            length = _poly_length(elem)
            if length < 100:
                continue
            result.append(elem)
    return result


# =============================================
# =============================================
# BEAM DETECTION – Cluster → Merge → Pair → Validate
# =============================================
class BeamAxis:
    """
    Trục dầm phát hiện từ 2 cluster line song song.
    type      = 'beam_axis'
    points    = [start, end]  – 2 đầu location line
    width     = khoảng cách 2 cluster (mm) = bề rộng dầm
    beam_type = 'edge' | 'middle'
                edge   → location line = nét ngoài cùng (trimmed)
                middle → location line = tim 2 cluster (trimmed)
    layer     = layer của cluster ngoài (edge) hoặc cluster A (middle)
    """
    def __init__(self, start, end, width, layer='', beam_type='middle'):
        self.type       = 'beam_axis'
        self.points     = [start, end]
        self.start      = start
        self.end        = end
        self.width      = _round5(width)
        self.center     = None
        self.radius     = 0.0
        self.layer      = layer
        self.beam_type  = beam_type   # 'edge' or 'middle'
        self.h          = 0           # chiều cao dầm từ text CAD (0 = chưa có)
        self.text_label = None        # chuỗi 'bxh' parse được từ text gần nhất
        self.location_type_override = None  # None=auto | 0=Left | 1=Right | 2=Center | 3=Origin
        self.family_type_override   = None  # None=use group.FamilyType | str=per-element override

    @property
    def height(self):
        """Chiều dài trục dầm."""
        return _poly_length(self)


def _explode_to_segments(elements):
    """
    Tách tất cả line/polyline thành list segment [(p0,p1,layer), ...]
    """
    segs = []
    for elem in elements:
        if elem.type == 'circle':
            continue
        pts = getattr(elem, 'points', [])
        lyr = getattr(elem, 'layer', '')
        if len(pts) < 2:
            continue
        for i in range(len(pts) - 1):
            segs.append((pts[i], pts[i+1], lyr))
    return segs


def _seg_angle(p0, p1):
    """Góc đường thẳng trong [0, π)."""
    dx = p1[0] - p0[0]
    dy = p1[1] - p0[1]
    a  = math.atan2(dy, dx)
    if a < 0:
        a += math.pi
    if a >= math.pi:
        a -= math.pi
    return a


def _seg_length(p0, p1):
    return math.sqrt((p1[0]-p0[0])**2 + (p1[1]-p0[1])**2)


def _seg_midpoint(p0, p1):
    return ((p0[0]+p1[0])/2.0, (p0[1]+p1[1])/2.0)


def _perp_offset(p0, p1, ref):
    """
    Khoảng cách có dấu từ điểm ref đến đường thẳng qua p0-p1.
    """
    dx = p1[0] - p0[0]
    dy = p1[1] - p0[1]
    L  = math.sqrt(dx*dx + dy*dy)
    if L < 1e-6:
        return 0.0
    return (-dy * (ref[0]-p0[0]) + dx * (ref[1]-p0[1])) / L


def _axis_overlap(p0a, p1a, p0b, p1b, angle):
    """
    Tỷ lệ overlap chiếu trên trục chung.
    Trả về (overlap_ratio, overlap_length).
    """
    cos_a = math.cos(angle)
    sin_a = math.sin(angle)

    def _proj(p): return p[0]*cos_a + p[1]*sin_a

    pa0, pa1 = sorted([_proj(p0a), _proj(p1a)])
    pb0, pb1 = sorted([_proj(p0b), _proj(p1b)])

    overlap = max(0.0, min(pa1, pb1) - max(pa0, pb0))
    min_len  = min(pa1-pa0, pb1-pb0)
    if min_len < 1e-6:
        return 0.0, 0.0
    return overlap / min_len, overlap


def _merge_collinear_segs(segs, angle_tol=0.052, gap_tol=50.0):
    """
    Gom các segments collinear (cùng hướng + cùng offset) vào 1 segment dài.
    angle_tol ≈ 3°, gap_tol = 50 mm.
    """
    if not segs:
        return []

    merged = []
    used   = [False] * len(segs)

    def _same_line(i, j):
        p0a, p1a, _ = segs[i]
        p0b, p1b, _ = segs[j]
        ang_a = _seg_angle(p0a, p1a)
        ang_b = _seg_angle(p0b, p1b)
        da    = abs(ang_a - ang_b)
        if da > math.pi/2: da = math.pi - da
        if da > angle_tol:
            return False
        # khoảng cách 2 đường thẳng
        d = abs(_perp_offset(p0a, p1a, p0b))
        return d < gap_tol

    for i in range(len(segs)):
        if used[i]: continue
        group   = [i]
        used[i] = True
        for j in range(i+1, len(segs)):
            if not used[j] and _same_line(i, j):
                group.append(j)
                used[j] = True
        # Gom tất cả points trong group → min/max projection
        p0_ref, p1_ref, lyr = segs[group[0]]
        angle = _seg_angle(p0_ref, p1_ref)
        cos_a = math.cos(angle)
        sin_a = math.sin(angle)
        def _proj(p): return p[0]*cos_a + p[1]*sin_a
        all_pts = []
        for gi in group:
            all_pts.append(segs[gi][0])
            all_pts.append(segs[gi][1])
        projs  = [(_proj(p), p) for p in all_pts]
        projs.sort(key=lambda x: x[0])
        new_p0 = projs[0][1]
        new_p1 = projs[-1][1]
        merged.append((new_p0, new_p1, lyr))

    return merged


def _merge_segs_in_cluster(segs, p0_ref, ref_angle, merge_tol):
    """
    Merge các segment trong 1 offset-cluster thành các interval liên tục.
    Sort theo projection lên ref_angle, merge nếu overlap OR gap <= merge_tol.
    Trả về list of (proj_start, proj_end, layer) đã sort.
    """
    if not segs:
        return []
    cos_a     = math.cos(ref_angle)
    sin_a     = math.sin(ref_angle)
    ref_proj  = p0_ref[0]*cos_a + p0_ref[1]*sin_a

    intervals = []
    for (p0, p1, lyr) in segs:
        t0 = p0[0]*cos_a + p0[1]*sin_a - ref_proj
        t1 = p1[0]*cos_a + p1[1]*sin_a - ref_proj
        if t0 > t1:
            t0, t1 = t1, t0
        intervals.append((t0, t1, lyr))
    intervals.sort(key=lambda x: x[0])

    merged = []
    cur_s, cur_e, cur_lyr = intervals[0]
    for t0, t1, lyr in intervals[1:]:
        if t0 <= cur_e + merge_tol:      # overlap hoặc gap nhỏ → nối
            cur_e = max(cur_e, t1)
        else:
            merged.append((cur_s, cur_e, cur_lyr))
            cur_s, cur_e, cur_lyr = t0, t1, lyr
    merged.append((cur_s, cur_e, cur_lyr))
    return merged


def _union_range(intervals):
    """Tính (min_start, max_end) của tập intervals [(t_s, t_e, lyr), ...]."""
    return min(iv[0] for iv in intervals), max(iv[1] for iv in intervals)


def detect_beams_from_lines(elements,
                             angle_tol_deg=5.0,
                             offset_tol=30.0,
                             merge_tol=100.0,
                             min_d=100.0,
                             max_d=1000.0,
                             min_overlap_len=200.0):
    """
    Phát hiện dầm theo đúng thứ tự:
      1. Explode → segments (lọc ngắn)
      2. Direction clustering (angle_tol_deg)
      3. Với mỗi direction group:
         3.1 Cluster theo offset vuông góc (offset_tol)
         3.2 Merge trong từng offset-cluster (merge_tol)
         3.3 Sort cluster theo offset tăng dần
         3.4 Pair chỉ cluster kề nhau (adjacent)
         3.5 Kiểm tra min_d <= d <= max_d
         3.6 Tính overlap; bỏ nếu overlap_len < min_overlap_len
         3.7 Xác định dầm biên / giữa → tạo BeamAxis
    """
    angle_tol = math.radians(angle_tol_deg)

    # ── Step 1: Explode ──────────────────────────────────────────────────
    segs = _explode_to_segments(elements)
    segs = [(p0, p1, lyr) for p0, p1, lyr in segs if _seg_length(p0, p1) >= 50]
    if not segs:
        return []

    # ── Step 2: Direction clustering ─────────────────────────────────────
    dir_clusters = []   # list of [ref_angle, [segs]]

    for seg in segs:
        p0, p1, lyr = seg
        a = _seg_angle(p0, p1)
        placed = False
        for dc in dir_clusters:
            da = abs(a - dc[0])
            if da > math.pi/2: da = math.pi - da
            if da <= angle_tol:
                dc[1].append(seg)
                placed = True
                break
        if not placed:
            dir_clusters.append([a, [seg]])

    beam_axes = []

    for ref_angle, dir_segs in dir_clusters:
        if len(dir_segs) < 2:
            continue

        # Điểm gốc tham chiếu để đo offset và projection
        p0_ref = dir_segs[0][0]
        cos_a  = math.cos(ref_angle)
        sin_a  = math.sin(ref_angle)
        # p1_ref dùng cho _perp_offset
        p1_ref = (p0_ref[0] + cos_a, p0_ref[1] + sin_a)

        # ── Step 3.1: Cluster theo offset ────────────────────────────────
        # offset_clusters: list of [mean_offset, [segs]]
        offset_clusters = []
        for seg in dir_segs:
            mid = _seg_midpoint(seg[0], seg[1])
            off = _perp_offset(p0_ref, p1_ref, mid)
            placed = False
            for oc in offset_clusters:
                if abs(off - oc[0]) <= offset_tol:
                    oc[1].append(seg)
                    # cập nhật mean offset
                    oc[0] = sum(
                        _perp_offset(p0_ref, p1_ref, _seg_midpoint(s[0], s[1]))
                        for s in oc[1]
                    ) / len(oc[1])
                    placed = True
                    break
            if not placed:
                offset_clusters.append([off, [seg]])

        offset_clusters.sort(key=lambda x: x[0])
        n_clusters = len(offset_clusters)
        if n_clusters < 2:
            continue

        # ── Step 3.2: Merge trong từng offset-cluster ────────────────────
        # merged_clusters: list of (mean_offset, [(t_start, t_end, lyr)])
        merged_clusters = []
        for mean_off, oc_segs in offset_clusters:
            intervals = _merge_segs_in_cluster(oc_segs, p0_ref, ref_angle, merge_tol)
            if intervals:
                lyr = intervals[0][2]
                merged_clusters.append((mean_off, intervals, lyr))

        if len(merged_clusters) < 2:
            continue

        n_mc = len(merged_clusters)

        # ── Step 3.3 + 3.4: Sort (đã sort) → Pair kề nhau ──────────────
        for gi in range(n_mc - 1):
            off_a, ivs_a, lyr_a = merged_clusters[gi]
            off_b, ivs_b, lyr_b = merged_clusters[gi + 1]

            # ── Step 3.5: Kiểm tra khoảng cách ──────────────────────────
            d = abs(off_b - off_a)
            if d < min_d or d > max_d:
                continue

            # ── Step 3.6: Tính overlap ───────────────────────────────────
            a_s, a_e = _union_range(ivs_a)
            b_s, b_e = _union_range(ivs_b)
            ov_start = max(a_s, b_s)
            ov_end   = min(a_e, b_e)
            ov_len   = ov_end - ov_start
            if ov_len < min_overlap_len:
                continue

            # ── Step 3.7: Xác định dầm biên / giữa ─────────────────────
            # Edge = một trong hai cluster là ngoài cùng của toàn nhóm
            is_edge = (gi == 0 or gi + 1 == n_mc - 1)
            beam_type = 'edge' if is_edge else 'middle'

            if beam_type == 'edge':
                # Location line = nét ngoài cùng (cluster xa tâm nhóm hơn)
                all_offsets   = [mc[0] for mc in merged_clusters]
                bbox_mid_off  = (min(all_offsets) + max(all_offsets)) / 2.0
                loc_off = off_a if abs(off_a - bbox_mid_off) > abs(off_b - bbox_mid_off) else off_b
            else:
                # Location line = tim giữa 2 cluster
                loc_off = (off_a + off_b) / 2.0

            # Chuyển (proj trên trục, offset vuông góc) → tọa độ thực
            ref_proj_base = p0_ref[0]*cos_a + p0_ref[1]*sin_a

            def _to_xy(t_proj, perp_off):
                base_x = p0_ref[0] + t_proj * cos_a
                base_y = p0_ref[1] + t_proj * sin_a
                return (base_x + (-sin_a) * perp_off,
                        base_y +   cos_a  * perp_off)

            ax_start = _to_xy(ov_start, loc_off)
            ax_end   = _to_xy(ov_end,   loc_off)

            layer = lyr_a or lyr_b
            beam_axes.append(BeamAxis(ax_start, ax_end, d, layer, beam_type))

    return beam_axes


# ── Legacy wrapper (giữ tương thích nếu có nơi gọi cũ) ──────────────────
def _detect_beams_legacy(elements,
                          angle_tol_deg=5.0,
                          offset_tol=50.0,
                          min_width=100.0,
                          max_width=1000.0,
                          min_overlap=0.60):
    """Wrapper chuyển sang API mới."""
    return detect_beams_from_lines(
        elements,
        angle_tol_deg   = angle_tol_deg,
        offset_tol      = offset_tol,
        merge_tol       = 100.0,
        min_d           = min_width,
        max_d           = max_width,
        min_overlap_len = 200.0,
    )


# ── DEAD CODE ─ kept for reference only ─────────────────────────────────
def _dead_detect_beams_old(elements,
                            angle_tol_deg=5.0,
                            offset_tol=50.0,
                            min_width=100.0,
                            max_width=1000.0,
                            min_overlap=0.60):
    """Old version – Merge TRƯỚC direction cluster (sai thứ tự, giữ tham khảo)."""
    angle_tol = math.radians(angle_tol_deg)
    segs = _explode_to_segments(elements)
    if not segs:
        return []
    # (BUG: merge global truoc khi cluster theo direction)
    segs = _merge_collinear_segs(segs, angle_tol=math.radians(3), gap_tol=50.0)
    dir_clusters = []
    dir_angles   = []
    for seg in segs:
        p0, p1, lyr = seg
        if _seg_length(p0, p1) < 50:
            continue
        a = _seg_angle(p0, p1)
        placed = False
        for ci, ca in enumerate(dir_angles):
            da = abs(a - ca)
            if da > math.pi/2: da = math.pi - da
            if da <= angle_tol:
                dir_clusters[ci].append(seg)
                placed = True
                break
        if not placed:
            dir_clusters.append([seg])
            dir_angles.append(a)
    beam_axes = []
    for ci, cluster in enumerate(dir_clusters):
        if len(cluster) < 2:
            continue
        ref_angle = dir_angles[ci]
        p0_ref, p1_ref, _ = cluster[0]
        seg_offsets = []
        for seg in cluster:
            mid = _seg_midpoint(seg[0], seg[1])
            off = _perp_offset(p0_ref, p1_ref, mid)
            seg_offsets.append((off, seg))
        seg_offsets.sort(key=lambda x: x[0])
        offset_groups = []
        for off, seg in seg_offsets:
            placed = False
            for og in offset_groups:
                if abs(off - og[0]) <= offset_tol:
                    og[1].append(seg)
                    og[0] = sum(_perp_offset(p0_ref, p1_ref, _seg_midpoint(s[0], s[1]))
                                for s in og[1]) / len(og[1])
                    placed = True
                    break
            if not placed:
                offset_groups.append([off, [seg]])
        offset_groups.sort(key=lambda x: x[0])
        if len(offset_groups) < 2:
            continue
        for gi in range(len(offset_groups) - 1):
            off_a, segs_a = offset_groups[gi]
            off_b, segs_b = offset_groups[gi+1]
            width = abs(off_b - off_a)
            if width < min_width or width > max_width:
                continue
            def _best_seg(sg_list):
                return max(sg_list, key=lambda s: _seg_length(s[0], s[1]))
            sa = _best_seg(segs_a)
            sb = _best_seg(segs_b)
            ovr, ovr_len = _axis_overlap(sa[0], sa[1], sb[0], sb[1], ref_angle)
            if ovr < min_overlap or ovr_len < 100:
                continue
            # (phần dựng BeamAxis cũ đã bị bỏ – xem detect_beams_from_lines mới)
            pass

    return beam_axes


# =============================================
# STEP 7 – GHÉP TEXT VÀO DẦM (beam axis)
# =============================================
def _pair_texts_with_beams(beam_axes, texts):
    """
    Ghép text kích thước gần nhất vào mỗi BeamAxis.
    texts: list of (content, x, y, layer)
    Kết quả: beam.h và beam.text_label được cập nhật in-place.
    """
    if not beam_axes or not texts:
        return
    for beam in beam_axes:
        mid = _seg_midpoint(beam.start, beam.end)
        best_dsq  = None
        best_dims = None
        for (content, tx, ty, _lyr) in texts:
            dims = _extract_beam_dims(content)
            if dims is None:
                continue
            dsq = _dist_sq(mid, (tx, ty))
            if best_dsq is None or dsq < best_dsq:
                best_dsq  = dsq
                best_dims = dims
        if best_dims is not None:
            beam.h          = best_dims[1]   # height từ text
            beam.text_label = u'{}x{}'.format(best_dims[0], best_dims[1])


# =============================================
# AXIS ALIGN – Làm tròn về hệ tọa độ trục
# =============================================
def align_elements_to_axis(elements, grid_elem):
    """
    Snap tọa độ location line (dầm) / tâm (cột, móng) về bội số 5mm
    trong hệ tọa độ địa phương của grid_elem.

    grid_elem : CadElement (line/polyline) – đường trục tham chiếu
    Trả về (elements_snapped, status):
        status = 'v'  → OK
        status = '!'  → grid_elem không hợp lệ
    """
    if grid_elem is None:
        return elements, '!'
    pts = getattr(grid_elem, 'points', [])
    if len(pts) < 2:
        return elements, '!'

    p0, p1 = pts[0], pts[-1]
    dx = p1[0] - p0[0]
    dy = p1[1] - p0[1]
    L  = math.sqrt(dx*dx + dy*dy)
    if L < 1e-6:
        return elements, '!'

    ux, uy = dx / L, dy / L          # trục u = dọc theo grid
    vx, vy = -uy, ux                  # trục v = vuông góc grid

    def _to_local(pt):
        qx, qy = pt[0] - p0[0], pt[1] - p0[1]
        return qx*ux + qy*uy, qx*vx + qy*vy

    def _to_world(u_loc, v_loc):
        return (p0[0] + u_loc*ux + v_loc*vx,
                p0[1] + u_loc*uy + v_loc*vy)

    def _snap_pt(pt):
        u_loc, v_loc = _to_local(pt)
        return _to_world(_round5(u_loc), _round5(v_loc))

    for elem in elements:
        if elem.type == 'beam_axis':
            elem.start  = _snap_pt(elem.start)
            elem.end    = _snap_pt(elem.end)
            elem.points = [elem.start, elem.end]

        elif elem.type == 'circle' and elem.center:
            new_c = _snap_pt(elem.center)
            dx_s  = new_c[0] - elem.center[0]
            dy_s  = new_c[1] - elem.center[1]
            elem.center = new_c
            if elem.points:
                elem.points = [(p[0]+dx_s, p[1]+dy_s) for p in elem.points]

        elif elem.type == 'polyline' and elem.points:
            xs = [p[0] for p in elem.points]
            ys = [p[1] for p in elem.points]
            cx = sum(xs) / len(xs)
            cy = sum(ys) / len(ys)
            new_cx, new_cy = _snap_pt((cx, cy))
            dx_s = new_cx - cx
            dy_s = new_cy - cy
            elem.points = [(p[0]+dx_s, p[1]+dy_s) for p in elem.points]

    return elements, 'v'


# =============================================
# ANALYZE ELEMENT (kích thước cụ thể)
# =============================================
def analyze_element(elem):
    """
    Trả về dict mô tả shape + kích thước:
    {'shape': 'REC'|'CIR'|'BEA'|'OTHER', 'label': str, 'w', 'h', 'dia'}
    """
    if elem.type == 'circle':
        dia = _round5(_mm(elem.radius * 2))
        return {'shape': 'CIR', 'label': 'CIR: {}'.format(dia), 'dia': dia}

    # Trục dầm (Method 1)
    if elem.type == 'beam_axis':
        w = getattr(elem, 'width',  0) or 0
        h = _round5(_poly_length(elem))   # chiều dài trục, không phải tiết diện
        return {'shape': 'BEA', 'label': 'BEA: {}x{}'.format(w, h), 'w': w, 'h': h}

    if elem.type in ('polyline', 'beam_line') and len(getattr(elem, 'points', [])) >= 2:
        pts = elem.points
        mn_x, mn_y, mx_x, mx_y = _bbox(pts)
        raw_w = mx_x - mn_x
        raw_h = mx_y - mn_y
        if raw_w < 1e-3 or raw_h < 1e-3:
            # Có thể là line nằm ngang/dọc → beam
            if elem.type == 'beam_line':
                w = getattr(elem, 'width', 0) or _round5(_mm(raw_w or raw_h))
                h = getattr(elem, 'height', 0) or w
                return {'shape': 'BEA', 'label': 'BEA: {}x{}'.format(w, h), 'w': w, 'h': h}
            return {'shape': 'OTHER', 'label': 'OTHER'}
        w = _round5(_mm(min(raw_w, raw_h)))
        h = _round5(_mm(max(raw_w, raw_h)))
        return {'shape': 'REC', 'label': 'REC: {}x{}'.format(w, h), 'w': w, 'h': h}

    return {'shape': 'OTHER', 'label': 'OTHER'}


def group_elements_by_label(elements):
    """
    Gom elements theo label (shape + kích thước).
    BeamAxis: group theo width (tiết diện), mỗi width = 1 group.
    Trả về list[dict] đã sắp xếp: REC trước, BEA sau, CIR, OTHER cuối.
    """
    groups = {}

    for elem in elements:
        # BeamAxis: group key = width x h (từ text label nếu có)
        if getattr(elem, 'type', '') == 'beam_axis':
            w   = getattr(elem, 'width', 0) or 0
            h   = getattr(elem, 'h', 0) or 0
            if h and h > 0:
                lbl = 'BEA: {}x{}'.format(w, h)
            else:
                lbl = 'BEA: {}x?'.format(w)
            if lbl not in groups:
                groups[lbl] = {'shape': 'BEA', 'label': lbl, 'w': w, 'h': h, 'elements': []}
            groups[lbl]['elements'].append(elem)
            continue

        info = analyze_element(elem)
        lbl  = info['label']
        if lbl == 'OTHER':
            continue
        if lbl not in groups:
            groups[lbl] = dict(info)
            groups[lbl]['elements'] = []
        groups[lbl]['elements'].append(elem)

    result = list(groups.values())

    def _sort_key(g):
        order = {'REC': 0, 'BEA': 1, 'CIR': 2, 'OTHER': 3}
        base  = order.get(g['shape'], 3)
        if g['shape'] == 'REC':
            return (base, g.get('w', 0), g.get('h', 0))
        if g['shape'] == 'BEA':
            return (base, g.get('w', 0), 0)
        if g['shape'] == 'CIR':
            return (base, g.get('dia', 0), 0)
        return (base, 0, 0)

    result.sort(key=_sort_key)
    return result


# =============================================
# BEAM ELEMENT (Line + Text)
# =============================================
def _parse_text_object(obj):
    """Đọc text content từ AcDbText hoặc AcDbMText."""
    try:
        name = obj.ObjectName.upper()
        if 'TEXT' in name:
            content = obj.TextString
            ins     = obj.InsertionPoint
            return (content, ins[0], ins[1])
    except Exception as ex:
        print('_parse_text_object error: {}'.format(ex))
    return None


def _extract_beam_dims(text_str):
    """Tìm WxH trong chuỗi text. Trả về (w, h) round5 hoặc None."""
    m = _re.search(r'(\d+)\s*[xX\xd7]\s*(\d+)', text_str)
    if not m:
        return None
    a = _round5(int(m.group(1)))
    b = _round5(int(m.group(2)))
    return (a, b)


def select_beam_elements_in_cad(doc):
    """
    Interactive (legacy workflow): chọn hỗn hợp line + text, tự động ghép cặp.
    Trả về list[CadBeamPair].

    NOTE: Trong workflow mới, lines được lấy từ detect_beams_from_lines()
    và texts được lấy từ extract_all_from_doc() theo Text Layer rule.
    Dùng pair_beams_with_text_layer() thay thế.
    """
    raw = _raw_select(
        doc,
        "\nChon cac doan thang dam VA text kich thuoc, roi nhan Enter...\n"
    )
    lines = []
    texts = []
    for obj in raw:
        try:
            parsed_text = _parse_text_object(obj)
            if parsed_text is not None:
                texts.append(parsed_text)
                continue
            elem = CadElement(obj)
            if elem.type in ('line', 'polyline') and len(elem.points) >= 2:
                lines.append(elem)
        except Exception as ex:
            print('select_beam_elements_in_cad skip: {}'.format(ex))

    return _pair_lines_with_texts(lines, texts)


def pair_beams_with_text_layer(beam_axes, all_texts, text_layer_name=''):
    """
    Workflow mới: ghép BeamAxis với texts từ Text Layer rule.
    beam_axes       : list[BeamAxis] – kết quả từ detect_beams_from_lines()
    all_texts       : list[(content, x, y, layer)] – từ extract_all_from_doc()
    text_layer_name : tên layer cần lọc (nếu rỗng → dùng tất cả texts)
    Cập nhật beam.h và beam.text_label in-place.
    """
    if text_layer_name:
        layer_up = text_layer_name.strip().upper()
        texts = [(c, x, y, lyr) for c, x, y, lyr in all_texts
                 if lyr.upper() == layer_up]
    else:
        texts = list(all_texts)
    _pair_texts_with_beams(beam_axes, texts)


def _pair_lines_with_texts(lines, texts):
    """Ghép mỗi line với text có kích thước gần nhất.
    texts có thể là 3-tuple (content, x, y) hoặc 4-tuple (content, x, y, layer).
    """
    result = []
    for line in lines:
        pts = line.points
        xs  = [p[0] for p in pts]
        ys  = [p[1] for p in pts]
        lc  = ((min(xs)+max(xs))/2.0, (min(ys)+max(ys))/2.0)

        best_text = None
        best_dsq  = None
        best_dims = None

        for t_entry in texts:
            txt, tx, ty = t_entry[0], t_entry[1], t_entry[2]
            dims = _extract_beam_dims(txt)
            if dims is None:
                continue
            dsq = _dist_sq(lc, (tx, ty))
            if best_dsq is None or dsq < best_dsq:
                best_dsq  = dsq
                best_text = txt
                best_dims = dims

        if best_dims is None:
            continue
        w, h = best_dims
        result.append(CadBeamPair(line, best_text, w, h))
    return result


def group_beam_pairs_by_label(pairs):
    """Gom CadBeamPair theo label 'BEA: WxH'."""
    groups = {}
    for pair in pairs:
        lbl = pair.label
        if lbl not in groups:
            groups[lbl] = {
                'label'   : lbl,
                'shape'   : 'BEA',
                'elements': [],
                'w'       : pair.width,
                'h'       : pair.height,
            }
        groups[lbl]['elements'].append(pair)
    return sorted(groups.values(), key=lambda g: (g['w'], g['h']))
