# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import Options, Solid, GeometryInstance, Transform

def extract_faces_with_transform(element):
    """
    Return List[(Face, Transform)]
    """
    results = []
    warnings = []

    opt = Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = False

    geo = element.get_Geometry(opt)
    if not geo:
        return results, warnings

    get_transform = getattr(element, 'GetTransform', None)
    base_transform = get_transform() if callable(get_transform) else Transform.Identity

    for obj in geo:
        if isinstance(obj, Solid) and obj.Faces.Size > 0:
            for f in obj.Faces:
                results.append((f, base_transform))
        elif isinstance(obj, GeometryInstance):
            inst_geo = obj.GetInstanceGeometry()
            tr = obj.Transform
            for g in inst_geo:
                if isinstance(g, Solid) and g.Faces.Size > 0:
                    for f in g.Faces:
                        results.append((f, tr))

    return results, warnings
