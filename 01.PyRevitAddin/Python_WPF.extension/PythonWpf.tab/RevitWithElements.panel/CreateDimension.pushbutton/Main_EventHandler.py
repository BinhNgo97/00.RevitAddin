# -*- coding: utf-8 -*-

from Autodesk.Revit.UI import IExternalEventHandler
from Autodesk.Revit.DB import Transaction, ReferenceArray
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

from geometry_extractors import extract_faces_with_transform
from face_filter import filter_faces_for_element
from face_grouper import group_faces_by_direction
from face_selection_policy import select_representative_faces_by_direction
from utility import get_bbox_center, create_dim_line_from_faces


# ------------------------------------------------------------
# Selection Filter
# ------------------------------------------------------------

class _ElementSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return True

    def AllowReference(self, reference, position):
        return False


# ------------------------------------------------------------
# Main Event Handler
# ------------------------------------------------------------

class MainEventHandler(IExternalEventHandler):
    """
    ExternalEvent Handler
    - Nhận action từ UI
    - Thực thi Revit API trong Execute
    """

    def __init__(self, window, viewmodel):
        self.window = window
        self.vm = viewmodel

        self.action = None
        self.message = None
        self.OnCompleted = None
        self.is_busy = False

    # --------------------------------------------------
    # IExternalEventHandler
    # --------------------------------------------------

    def GetName(self):
        return "Create Dimension External Event"

    def Execute(self, app):
        if self.is_busy:
            return

        self.is_busy = True
        success = False

        try:
            if self.action == "ELEMENT_REFPLAN":
                success = self._dim_element_refplan(app)
            elif self.action == "GRID_LEVEL":
                self.message = "GRID/LEVEL is not implemented yet"
                success = False
            else:
                self.message = "Unknown action: {}".format(self.action)

        except Exception as ex:
            self.message = str(ex)
            success = False

        finally:
            self.is_busy = False
            if self.OnCompleted:
                self.OnCompleted(success, self.action)

    # --------------------------------------------------
    # CORE LOGIC
    # --------------------------------------------------

    def _dim_element_refplan(self, app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView

        # --------------------------------------------------
        # 1. Pick elements (MULTI)
        # --------------------------------------------------
        try:
            picks = uidoc.Selection.PickObjects(
                ObjectType.Element,
                _ElementSelectionFilter(),
                "Pick elements to dimension"
            )
        except:
            self.message = "Selection cancelled"
            return False

        elements = [doc.GetElement(p.ElementId) for p in picks if p]

        if not elements:
            self.message = "No elements selected"
            return False

        # --------------------------------------------------
        # 2. Extract + filter faces
        # --------------------------------------------------
        all_faces = []

        for el in elements:
            faces, _ = extract_faces_with_transform(el)
            fds, _ = filter_faces_for_element(el, faces, view)

            if fds:
                all_faces.extend(fds)

        if len(all_faces) < 2:
            self.message = "Not enough valid faces"
            return False

        # --------------------------------------------------
        # 3. Group faces by direction
        # --------------------------------------------------
        direction_map = group_faces_by_direction(all_faces, view)

        if not direction_map:
            self.message = "No valid dimension directions"
            return False

        # --------------------------------------------------
        # 4. Face selection policy
        #   - 1 face / element / direction
        #   - faces SORTED theo direction
        # --------------------------------------------------
        direction_map = select_representative_faces_by_direction(
            direction_map,
            view
        )

        # --------------------------------------------------
        # 5. Bounding box center (for dim line)
        # --------------------------------------------------
        center = get_bbox_center(elements, view)
        if not center:
            self.message = "Cannot compute bounding box center"
            return False

        created_any = False

        # --------------------------------------------------
        # 6. CREATE DIMENSION
        #   - 1 line (biên ngoài)
        #   - ReferenceArray = ALL faces
        # --------------------------------------------------
        with Transaction(doc, "Auto Dimension Elements") as t:
            t.Start()

            for direction, faces in direction_map.items():
                if len(faces) < 2:
                    continue

                # DEBUG: Print info
                print("\n=== Direction Group ===")
                print("Direction: ({:.3f}, {:.3f}, {:.3f})".format(direction.X, direction.Y, direction.Z))
                print("Number of faces: {}".format(len(faces)))
                
                element_ids = set()
                for i, f in enumerate(faces):
                    o = f.origin
                    n = f.normal
                    elem_id = f.element_id
                    element_ids.add(str(elem_id))
                    dist_along = o.DotProduct(direction)
                    # Check if normal is same or opposite to direction
                    normal_dot = n.Normalize().DotProduct(direction)
                    print("  Face {}: ElementId={}, origin=({:.2f}, {:.2f}, {:.2f}), normal=({:.3f}, {:.3f}, {:.3f}), dist={:.2f}, normal_dot={:.3f}".format(
                        i, elem_id, o.X, o.Y, o.Z, n.X, n.Y, n.Z, dist_along, normal_dot))
                
                print("  Unique elements: {}".format(len(element_ids)))
                if len(element_ids) < 2:
                    print("  WARNING: Only {} unique element(s)! Should have 2+".format(len(element_ids)))

                # faces đã:
                # - 1 face / element
                # - sort theo direction

                # ---- LINE: BIÊN NGOÀI ----
                f_start = faces[0]
                f_end = faces[-1]

                # ---- REFERENCES: LẤY HẾT ----
                ra = ReferenceArray()
                for f in faces:
                    ra.Append(f.reference)

                dim_line = create_dim_line_from_faces(
                    view,
                    center,
                    direction,
                    f_start,
                    f_end,
                    faces=faces,
                    offset=5.0  # Increased offset for better visibility
                )

                if not dim_line:
                    print("  ERROR: Failed to create dim_line")
                    continue

                print("  Dim line: {} -> {}".format(
                    dim_line.GetEndPoint(0), dim_line.GetEndPoint(1)))

                doc.Create.NewDimension(view, dim_line, ra)
                created_any = True

            t.Commit()

        self.message = (
            "Dimension created"
            if created_any
            else "No valid dimension created"
        )

        return created_any
