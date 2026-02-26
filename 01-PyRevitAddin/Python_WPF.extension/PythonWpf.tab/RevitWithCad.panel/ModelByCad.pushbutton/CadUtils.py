# -*- coding: utf-8 -*-
"""
CadUtils.py - Tiện ích kết nối và lấy dữ liệu từ AutoCAD COM.

Exports:
    get_acad_doc()               - Kết nối tới AutoCAD, trả về ActiveDocument
    select_elements_in_cad(doc)  - Cho người dùng chọn elements (polyline, line, circle,...)
    select_grid_in_cad(doc)      - Cho người dùng chọn 2 đường trục
    CadElement                   - Dataclass lưu thông tin hình học 1 element CAD
"""
import math
import clr
from System.Runtime.InteropServices import Marshal
from System import Type, Activator


# =============================================
# DATA CLASS
# =============================================
class CadElement:
    """
    Lưu thông tin hình học đã cache của 1 đối tượng CAD.
    type : 'polyline' | 'line' | 'circle' | 'arc' | 'unknown'
    points : list of (x, y)   – dùng cho polyline / line
    center : (x, y)           – dùng cho circle / arc
    radius : float            – dùng cho circle / arc
    start_angle / end_angle   – dùng cho arc (radian)
    obj    : tham chiếu COM gốc (để đổi màu/layer nếu cần)
    """
    def __init__(self, obj):
        self.obj = obj
        self.points = []
        self.center = None
        self.radius = 0.0
        self.start_angle = 0.0
        self.end_angle = 0.0
        self.type = 'unknown'
        self._parse(obj)

    def _parse(self, obj):
        try:
            name = obj.ObjectName.upper()

            if 'CIRCLE' in name:
                self.type = 'circle'
                c = obj.Center
                self.center = (c[0], c[1])
                self.radius = obj.Radius

            elif 'ARC' in name:
                self.type = 'arc'
                c = obj.Center
                self.center = (c[0], c[1])
                self.radius = obj.Radius
                self.start_angle = obj.StartAngle
                self.end_angle = obj.EndAngle

            elif 'LINE' in name:
                # AcDbLine      → ObjectName = "AcDbLine"   → uppercased "ACDBLINE"
                # AcDbLWPolyline→ ObjectName = "AcDbPolyline" → uppercased "ACDBLWPOLYLINE"
                # AcDb2dPolyline→ uppercased "ACDB2DPOLYLINE"
                if 'POLYLINE' not in name:
                    # Đây là LINE đơn giản – không có thuộc tính Coordinates
                    sp = obj.StartPoint
                    ep = obj.EndPoint
                    self.points = [(sp[0], sp[1]), (ep[0], ep[1])]
                    self.type = 'line'
                else:
                    coords = obj.Coordinates
                    # Phát hiện step an toàn: LWPolyline lưu XY (step=2),
                    # 2D/3D Polyline lưu XYZ (step=3)
                    # Ưu tiên kiểm tra chia hết; tên object không đủ tin cậy
                    if len(coords) % 2 == 0:
                        step = 2
                    else:
                        step = 3
                    n = len(coords) // step
                    self.points = [(coords[i*step], coords[i*step+1]) for i in range(n)]
                    self.type = 'polyline'
                    # Nếu polyline đã kín (Closed=True) nhưng điểm cuối ≠ điểm đầu
                    # thì thêm điểm đầu vào cuối để vòng được đóng kín
                    try:
                        if obj.Closed and len(self.points) >= 2:
                            p0, pn = self.points[0], self.points[-1]
                            if (p0[0] - pn[0])**2 + (p0[1] - pn[1])**2 > 1e-6:
                                self.points.append(p0)
                    except Exception:
                        pass
            else:
                self.type = 'unknown'
        except Exception as ex:
            print('CadElement parse error: {}'.format(ex))
            self.type = 'unknown'


# =============================================
# CONNECTION
# =============================================
def get_acad_doc():
    """Kết nối an toàn tới AutoCAD, trả về ActiveDocument hoặc None."""
    try:
        acad = Marshal.GetActiveObject("AutoCAD.Application")
        return acad.ActiveDocument
    except:
        try:
            acad_type = Type.GetTypeFromProgID("AutoCAD.Application")
            acad = Activator.CreateInstance(acad_type)
            return acad.ActiveDocument
        except Exception as ex:
            print("Khong the ket noi AutoCAD: {0}".format(ex))
            return None


def _raw_select(doc, prompt_msg):
    """Hàm nội bộ: mở SelectionSet, cho người dùng chọn, trả về list COM objects."""
    raw = []
    ssa_name = "PySel_ModelByCad"
    try:
        doc.Utility.Prompt(prompt_msg)
        try:
            doc.SelectionSets.Item(ssa_name).Delete()
        except:
            pass
        ssa = doc.SelectionSets.Add(ssa_name)
        ssa.SelectOnScreen()
        for i in range(ssa.Count):
            raw.append(ssa.Item(i))
        ssa.Delete()
    except Exception as ex:
        print("Lỗi chọn đối tượng: {}".format(ex))
    return raw


# =============================================
# SELECT FUNCTIONS
# =============================================
def select_elements_in_cad(doc):
    """
    Cho người dùng chọn tự do các đối tượng (polyline, line, circle, arc...).
    Trả về list[CadElement] đã parse geometry.
    """
    raw = _raw_select(doc, "\nChọn các elements (polyline, line, circle...) và nhấn Enter...\n")
    result = []
    for obj in raw:
        try:
            elem = CadElement(obj)
            if elem.type != 'unknown':
                result.append(elem)
        except Exception as ex:
            print("Bỏ qua object lỗi: {}".format(ex))
    return result


def select_grid_in_cad(doc):
    """
    Cho người dùng chọn 1 đoạn thẳng tham chiếu trong CAD.
    - points[0] = StartPoint  → điểm gốc transform
    - points[-1] = EndPoint   → xác định phương chiều tham chiếu
    Trả về list[CadElement] với đúng 1 phần tử kiểu line.
    """
    raw = _raw_select(doc, "\nChọn ĐÚNG 1 đường tham chiếu trong CAD rồi nhấn Enter...\n")
    result = []
    for obj in raw:
        try:
            elem = CadElement(obj)
            if elem.type in ('line', 'polyline') and len(result) < 1:
                result.append(elem)
        except Exception as ex:
            print("Bỏ qua object lỗi: {}".format(ex))

    if len(result) != 1:
        print("Cảnh báo: cần chọn đúng 1 đường tham chiếu, hiện có {}.".format(len(result)))
    return result


# =============================================
# MERGE LINES → CLOSED POLYLINES
# =============================================
class _MergedPolyline:
    """Pseudo-CadElement đại diện cho polyline được gom từ nhiều line/segment."""
    def __init__(self, points):
        self.obj = None
        self.type = 'polyline'
        self.points = points
        self.center = None
        self.radius = 0.0
        self.start_angle = 0.0
        self.end_angle = 0.0


def merge_lines_to_closed_polylines(elements, tolerance=1.0):
    """
    Gom các line/polyline có endpoints chung thành các polyline kín (closed loops).

    Thuật toán:
      1. Tách tất cả segment (cặp điểm liên tiếp) từ mọi line/polyline.
         Nếu element đã là polyline kín (first ≈ last) thì giữ nguyên.
      2. Với các line/segment còn lại: xây chuỗi (chain) bằng cách nối
         endpoint trùng (trong tolerance) theo cả hai hướng (tail + head).
      3. Nếu đầu ≈ cuối → polyline kín; snap điểm cuối về đúng điểm đầu.

    Args:
        elements : list[CadElement | _MergedPolyline]
        tolerance: khoảng cách max để coi 2 điểm là cùng nhau (CAD units)

    Returns:
        list gồm _MergedPolyline cho mỗi chain + các CadElement khác giữ nguyên
    """
    tol_sq = tolerance ** 2

    def _dsq(a, b):
        return (a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2

    def _near(a, b):
        return _dsq(a, b) <= tol_sq

    line_elems  = [e for e in elements if e.type in ('line', 'polyline')]
    other_elems = [e for e in elements if e.type not in ('line', 'polyline')]

    if not line_elems:
        return elements

    # Tách segment. Polyline kín (first≈last hoặc >=3 pts đã kín) → giữ nguyên.
    already_closed = []
    segments = []          # list of (pt_start, pt_end)

    for elem in line_elems:
        pts = elem.points
        if len(pts) < 2:
            continue
        if len(pts) >= 3 and _near(pts[0], pts[-1]):
            # Đã kín, không cần gom thêm
            already_closed.append(_MergedPolyline(list(pts)))
        else:
            for i in range(len(pts) - 1):
                segments.append((pts[i], pts[i + 1]))

    if not segments:
        return already_closed + other_elems

    used = [False] * len(segments)

    def _find_next(pt):
        """
        Tìm segment chưa dùng có một đầu gần pt.
        Trả về (idx, new_pt) – new_pt là đầu còn lại của segment.
        """
        for i, (sa, sb) in enumerate(segments):
            if used[i]:
                continue
            if _near(pt, sa):
                return i, sb
            if _near(pt, sb):
                return i, sa
        return None, None

    chains = []

    for start_idx in range(len(segments)):
        if used[start_idx]:
            continue

        sa0, sb0 = segments[start_idx]
        chain = [sa0, sb0]
        used[start_idx] = True

        # Kéo dài về phía cuối (tail)
        while True:
            idx, nxt = _find_next(chain[-1])
            if idx is None:
                break
            used[idx] = True
            chain.append(nxt)
            # Nếu đã kín thì dừng sớm
            if _near(chain[0], chain[-1]):
                chain[-1] = chain[0]
                break
            # Giới hạn độ dài để tránh vòng lặp vô hạn
            if len(chain) > len(segments) + 2:
                break

        # Kéo dài về phía đầu (head) nếu chưa kín
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

    merged = already_closed + [_MergedPolyline(ch) for ch in chains]
    return merged + other_elems


# =============================================
# BEAM ELEMENT (Line + Text pair)
# =============================================
import re as _re


class CadBeamPair:
    """
    Đại diện cho 1 dầm CAD = 1 line/polyline + 1 text gần nhất.
    Attributes:
        start   : (x, y)  – điểm đầu location line (mm)
        end     : (x, y)  – điểm cuối location line (mm)
        width   : int     – chiều rộng dầm (mm, round5), từ text
        height  : int     – chiều cao dầm  (mm, round5), từ text
        label   : str     – VD "BEA: 200x400"
        raw_text: str     – chuỗi text gốc trong CAD
        line_elem: CadElement – element đường thẳng gốc
    """
    def __init__(self, line_elem, text_str, width, height):
        self.line_elem = line_elem
        self.start     = line_elem.points[0]
        self.end       = line_elem.points[-1]
        self.raw_text  = text_str
        self.width     = width
        self.height    = height
        self.label     = 'BEA: {}x{}'.format(width, height)
        # type alias để drawing helpers nhận ra
        self.type      = 'beam_line'
        self.points    = line_elem.points   # cho bbox + preview


def _parse_text_object(obj):
    """
    Đọc nội dung text từ AcDbText hoặc AcDbMText (COM).
    Trả về (text_content_str, center_x, center_y) hoặc None nếu không đọc được.
    """
    try:
        name = obj.ObjectName.upper()
        if 'MTEXT' in name:
            content = obj.TextString
            ins     = obj.InsertionPoint
            return (content, ins[0], ins[1])
        elif 'TEXT' in name:
            content = obj.TextString
            ins     = obj.InsertionPoint
            return (content, ins[0], ins[1])
    except Exception as ex:
        print('_parse_text_object error: {}'.format(ex))
    return None


def _extract_beam_dims(text_str):
    """
    Tách kích thước WxH từ chuỗi text bất kỳ.
    Tìm pattern: (số)(x)(số), VD "GB1(200x400)" → (200, 400).
    Trả về (width_mm, height_mm) đã round5, hoặc None.
    """
    m = _re.search(r'(\d+)\s*[xX×]\s*(\d+)', text_str)
    if not m:
        return None
    a = _round5(int(m.group(1)))
    b = _round5(int(m.group(2)))
    # width = chiều nhỏ hơn, height = chiều lớn hơn (tùy convention)
    return (a, b)


def _dist_sq(p1, p2):
    return (p1[0] - p2[0])**2 + (p1[1] - p2[1])**2


def select_beam_elements_in_cad(doc):
    """
    Cho người dùng chọn hỗn hợp: các đoạn thẳng dầm + text kích thước.
    Tự động ghép mỗi line với text có tâm gần nhất.
    Trả về list[CadBeamPair] – mỗi phần tử là 1 dầm hợp lệ.
    """
    raw = _raw_select(
        doc,
        "\nChọn các đoạn thẳng dầm VÀ text kích thước, rồi nhấn Enter...\n"
    )

    lines = []
    texts = []   # list of (text_str, cx, cy)

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

    if not lines:
        print('Không tìm thấy đường thẳng nào trong selection.')
        return []

    # Tính tâm của mỗi line (midpoint)
    def _line_center(elem):
        xs = [p[0] for p in elem.points]
        ys = [p[1] for p in elem.points]
        return ((min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0)

    result = []

    for line in lines:
        lc = _line_center(line)

        # Tìm text gần nhất có chứa kích thước WxH
        best_text  = None
        best_dsq   = None
        best_dims  = None

        for (txt, tx, ty) in texts:
            dims = _extract_beam_dims(txt)
            if dims is None:
                continue
            dsq = _dist_sq(lc, (tx, ty))
            if best_dsq is None or dsq < best_dsq:
                best_dsq  = dsq
                best_text = txt
                best_dims = dims

        if best_dims is None:
            print('Dòng line không tìm được text kích thước, bỏ qua.')
            continue

        w, h = best_dims
        result.append(CadBeamPair(line, best_text, w, h))

    print('select_beam_elements_in_cad: {} lines, {} texts → {} pairs'.format(
        len(lines), len(texts), len(result)
    ))
    return result


def group_beam_pairs_by_label(pairs):
    """
    Gom CadBeamPair theo label (BEA: WxH).
    Trả về list of dict [{
        'label'   : str,
        'shape'   : 'BEA',
        'elements': list[CadBeamPair],
        'w'       : int,
        'h'       : int,
    }], sắp xếp theo kích thước tăng dần.
    """
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

    result = sorted(groups.values(), key=lambda g: (g['w'], g['h']))
    return result


# =============================================
# GEOMETRY ANALYSIS
# =============================================
def _round5(value):
    """Làm tròn đến bội số của 5 gần nhất. VD: 201→200, 204→205."""
    return int(round(value / 5.0) * 5)


def _mm(cad_units):
    """Chuyển đổi đơn vị CAD (mm) sang mm nguyên. Giữ nguyên nếu đã là mm."""
    return cad_units  # CAD thường đơn vị mm; điều chỉnh nếu cần


def analyze_element(elem):
    """
    Phân tích geometry của 1 element và trả về dict mô tả:
      {
        'shape'  : 'REC' | 'CIR' | 'OTHER',
        'label'  : str  ('REC: 100x200', 'CIR: 300', 'OTHER'),
        'w'      : float (mm, đã làm tròn 5) – cho REC
        'h'      : float (mm, đã làm tròn 5) – cho REC
        'dia'    : float (mm, đã làm tròn 5) – cho CIR
      }
    """
    if elem.type == 'circle':
        dia = _round5(_mm(elem.radius * 2))
        return {'shape': 'CIR', 'label': 'CIR: {}'.format(dia), 'dia': dia}

    if elem.type in ('polyline',) and len(elem.points) >= 3:
        pts = elem.points
        # Lấy bounding box của polyline
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        raw_w = max(xs) - min(xs)
        raw_h = max(ys) - min(ys)
        if raw_w < 1e-3 or raw_h < 1e-3:
            return {'shape': 'OTHER', 'label': 'OTHER'}
        w = _round5(_mm(raw_w))
        h = _round5(_mm(raw_h))
        # Kích thước nhỏ trước, lớn sau
        dim1, dim2 = (w, h) if w <= h else (h, w)
        return {'shape': 'REC', 'label': 'REC: {}x{}'.format(dim1, dim2), 'w': dim1, 'h': dim2}

    return {'shape': 'OTHER', 'label': 'OTHER'}


def group_elements_by_label(elements):
    """
    Gom các element có cùng label (shape + kích thước) thành nhóm.

    Returns:
        list of dict [{
            'label'     : str,
            'shape'     : str,
            'elements'  : list[CadElement/_MergedPolyline],
            'w'/'h'/'dia': float  (tuỳ shape)
        }]
        Sắp xếp: REC trước, CIR sau, OTHER cuối; cùng loại theo kích thước tăng dần.
    """
    groups = {}   # label → dict

    for elem in elements:
        info = analyze_element(elem)
        lbl = info['label']
        if lbl == 'OTHER':
            continue
        if lbl not in groups:
            groups[lbl] = dict(info)
            groups[lbl]['elements'] = []
        groups[lbl]['elements'].append(elem)

    result = list(groups.values())

    def _sort_key(g):
        order = {'REC': 0, 'CIR': 1, 'OTHER': 2}
        base = order.get(g['shape'], 2)
        if g['shape'] == 'REC':
            return (base, g.get('w', 0), g.get('h', 0))
        if g['shape'] == 'CIR':
            return (base, g.get('dia', 0), 0)
        return (base, 0, 0)

    result.sort(key=_sort_key)
    return result
