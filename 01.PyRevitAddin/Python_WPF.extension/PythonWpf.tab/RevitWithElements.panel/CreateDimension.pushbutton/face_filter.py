# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import XYZ, PlanarFace
from face_data import GroupFaceData

def filter_faces_for_element(element, faces_with_transform, view):
    """
    Filter valid faces for dimensioning.
    Apply transform to face normals (WORLD SPACE).
    """

    face_data_list = []
    warnings = []

    view_dir = view.ViewDirection.Normalize()

    for face, transform in faces_with_transform:
        if not isinstance(face, PlanarFace):
            continue

        if face.Reference is None:
            continue

        # APPLY TRANSFORM TO NORMAL (CRITICAL FIX)
        normal = transform.OfVector(face.FaceNormal).Normalize()

        # Skip face parallel to view direction (cannot dimension)
        if abs(normal.DotProduct(view_dir)) > 0.99:
            continue

        # Face center in world
        origin = transform.OfPoint(face.Origin)

        face_data_list.append(
            GroupFaceData(
                element_id=element.Id,
                planar_face=face,
                reference=face.Reference,
                normal=normal,
                origin=origin
            )
        )

    return face_data_list, warnings
