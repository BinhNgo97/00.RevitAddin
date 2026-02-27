# -*- coding: utf-8 -*-
ANGLE_DOT_TOL = 0.98

def _is_parallel(v1, v2):
    """Check if vectors are parallel (same or opposite direction)."""
    return abs(v1.Normalize().DotProduct(v2.Normalize())) >= ANGLE_DOT_TOL


def group_faces_by_direction(face_data_list, view):
    """
    Group faces by measurement direction.
    
    CRITICAL: For dimension between elements, we need OPPOSING faces
    (one pointing +direction, one pointing -direction)
    """
    view_dir = view.ViewDirection.Normalize()
    groups = {}

    for fd in face_data_list:
        # Measurement direction for linear dimension is the face normal
        # projected onto the view plane.
        n = fd.normal.Normalize()
        measure_dir = n - view_dir.Multiply(n.DotProduct(view_dir))
        if measure_dir.IsZeroLength():
            continue

        measure_dir = measure_dir.Normalize()

        # Find existing group with parallel direction (same or opposite)
        key_dir = None
        for d in groups.keys():
            if _is_parallel(d, measure_dir):
                key_dir = d
                break

        if key_dir:
            groups[key_dir].append(fd)
        else:
            groups[measure_dir] = [fd]

    return groups
