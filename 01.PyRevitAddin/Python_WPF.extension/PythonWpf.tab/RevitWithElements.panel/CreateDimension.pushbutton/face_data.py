# -*- coding: utf-8 -*-
# face_data.py
from Autodesk.Revit.DB import XYZ, PlanarFace

class GroupFaceData(object):
    def __init__(
        self,
        element_id,
        planar_face=None,
        face=None,
        reference=None,
        normal=None,
        origin=None,
    ):
        self.element_id = element_id          # ElementId

        # Backward-compatible inputs:
        # - old: (element_id, planar_face)
        # - new: keyword args (face/reference/normal/origin)
        if planar_face is None:
            planar_face = face

        self.face = planar_face               # PlanarFace
        self.reference = reference or (planar_face.Reference if planar_face else None)
        self.normal = normal or (planar_face.FaceNormal.Normalize() if planar_face else None)
        self.origin = origin or (planar_face.Origin if planar_face else None)

        self._distance_cache = {}

    def distance_along(self, direction):
        # Autodesk.Revit.DB.XYZ hashing can be implementation-dependent in IronPython.
        # Use a stable numeric key to avoid cache misses/errors.
        key = (
            round(direction.X, 9),
            round(direction.Y, 9),
            round(direction.Z, 9),
        )
        if key not in self._distance_cache:
            self._distance_cache[key] = self.origin.DotProduct(direction)
        return self._distance_cache[key]
