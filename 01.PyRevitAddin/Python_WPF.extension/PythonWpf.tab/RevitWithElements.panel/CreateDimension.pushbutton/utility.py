# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import XYZ, Line


# ------------------------------------------------------------
# BOUNDING BOX CENTER
# ------------------------------------------------------------

def get_bbox_center(elements, view):
    """
    Compute bounding box center of multiple elements in view.
    """

    bbox = None

    for el in elements:
        bb = el.get_BoundingBox(view)
        if not bb:
            continue

        if bbox is None:
            bbox = bb
        else:
            bbox.Min = XYZ(
                min(bbox.Min.X, bb.Min.X),
                min(bbox.Min.Y, bb.Min.Y),
                min(bbox.Min.Z, bb.Min.Z)
            )
            bbox.Max = XYZ(
                max(bbox.Max.X, bb.Max.X),
                max(bbox.Max.Y, bb.Max.Y),
                max(bbox.Max.Z, bb.Max.Z)
            )

    if not bbox:
        return None

    return (bbox.Min + bbox.Max) * 0.5


# ------------------------------------------------------------
# DIMENSION LINE CREATION (FIXED)
# ------------------------------------------------------------

def create_dim_line_from_faces(view, center, direction, face1, face2, faces=None, offset=1.0):
    """
    Create dimension line for rotated elements.
    
    Key fix:
    - Use ACTUAL face normal for distance calculation (not projected)
    - Line still parallel to measurement direction (projected normal)
    """

    if not faces or len(faces) < 2:
        return None

    # Get first face normal (actual, not projected)
    face_normal = getattr(face1, 'normal', None)
    if not face_normal:
        return None
    face_normal = face_normal.Normalize()

    # Get directions
    view_dir = view.ViewDirection.Normalize()
    measure_dir = direction.Normalize()  # Projected normal (for line direction)
    
    # Perpendicular to measurement (for offset positioning)
    perp_dir = view_dir.CrossProduct(measure_dir).Normalize()

    # Collect all face origins
    origins = []
    for f in faces:
        o = getattr(f, 'origin', None)
        if o:
            origins.append(o)
    
    if len(origins) < 2:
        return None

    # print("  view_dir: ({:.3f}, {:.3f}, {:.3f})".format(view_dir.X, view_dir.Y, view_dir.Z))
    # print("  face_normal (actual): ({:.3f}, {:.3f}, {:.3f})".format(face_normal.X, face_normal.Y, face_normal.Z))
    # print("  measure_dir (projected): ({:.3f}, {:.3f}, {:.3f})".format(measure_dir.X, measure_dir.Y, measure_dir.Z))
    # print("  perp_dir: ({:.3f}, {:.3f}, {:.3f})".format(perp_dir.X, perp_dir.Y, perp_dir.Z))

    # CRITICAL FIX: Use ACTUAL face normal for distance calculation
    distances = [o.DotProduct(face_normal) for o in origins]
    d_min = min(distances)
    d_max = max(distances)
    d_mid = (d_min + d_max) * 0.5
    
    actual_distance = abs(d_max - d_min)
    # print("  Actual distance (along face normal): {:.2f}".format(actual_distance))
    
    # Extend line slightly beyond faces (along measure_dir for line extent)
    measure_distances = [o.DotProduct(measure_dir) for o in origins]
    m_min = min(measure_distances)
    m_max = max(measure_distances)
    margin = (m_max - m_min) * 0.15 + 1.0
    line_start_dist = m_min - margin
    line_end_dist = m_max + margin

    # Find average position perpendicular to measurement
    perp_positions = [o.DotProduct(perp_dir) for o in origins]
    perp_avg = sum(perp_positions) / len(perp_positions)
    
    # Offset position perpendicular to measurement
    perp_position = perp_avg + offset

    # Build two points along measurement direction
    ref = origins[0]
    
    # Start point
    start = ref
    start = start + measure_dir.Multiply(line_start_dist - start.DotProduct(measure_dir))
    start = start + perp_dir.Multiply(perp_position - start.DotProduct(perp_dir))
    
    # End point
    end = ref
    end = end + measure_dir.Multiply(line_end_dist - end.DotProduct(measure_dir))
    end = end + perp_dir.Multiply(perp_position - end.DotProduct(perp_dir))

    # Keep the dimension line in the view plane to avoid "line not in plane" errors.
    # Project start/end onto the plane defined by (center, view_dir).
    if center is not None:
        plane_d = center.DotProduct(view_dir)
        start = start + view_dir.Multiply(plane_d - start.DotProduct(view_dir))
        end = end + view_dir.Multiply(plane_d - end.DotProduct(view_dir))

    return Line.CreateBound(start, end)
