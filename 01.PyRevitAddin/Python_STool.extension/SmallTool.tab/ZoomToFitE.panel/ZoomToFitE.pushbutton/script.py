# -*- coding: utf-8 -*-
# Zoom to Fit Selected Elements in Active View (pyRevit) - Auto Select if None

from Autodesk.Revit.DB import XYZ, View3D
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit

uidoc = revit.uidoc
doc = revit.doc
active_view = uidoc.ActiveView

# Lấy các Element đang được chọn
selection_ids = uidoc.Selection.GetElementIds()

# Nếu chưa chọn gì thì yêu cầu người dùng chọn
if not selection_ids:
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Chọn các đối tượng để zoom")
        selection_ids = [r.ElementId for r in refs]
    except:
        print("⚠ Không có đối tượng nào được chọn.")
        selection_ids = []

if selection_ids:
    selected_elements = [doc.GetElement(eid) for eid in selection_ids]
    valid_boxes = []

    for el in selected_elements:
        try:
            # Nếu là view 3D → dùng bounding box toàn mô hình
            if isinstance(active_view, View3D):
                bbox = el.get_BoundingBox(None)
            else:
                bbox = el.get_BoundingBox(active_view)

            if bbox:
                # Loại bỏ bounding box có tọa độ vô hạn
                if all(abs(v) != float('inf') for v in [bbox.Min.X, bbox.Min.Y, bbox.Max.X, bbox.Max.Y]):
                    valid_boxes.append(bbox)
        except:
            pass

    if not valid_boxes:
        print("⚠ Không tìm thấy BoundingBox hợp lệ trong View hiện tại.")
    else:
        min_pt = XYZ(min(b.Min.X for b in valid_boxes),
                     min(b.Min.Y for b in valid_boxes),
                     min(b.Min.Z for b in valid_boxes))
        max_pt = XYZ(max(b.Max.X for b in valid_boxes),
                     max(b.Max.Y for b in valid_boxes),
                     max(b.Max.Z for b in valid_boxes))

        for view in uidoc.GetOpenUIViews():
            if view.ViewId == active_view.Id:
                view.ZoomAndCenterRectangle(min_pt, max_pt)
                break
else:
    print("⚠ Không thể thực hiện lệnh vì không có đối tượng nào được chọn.")
