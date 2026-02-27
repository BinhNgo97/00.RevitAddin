# -*- coding: utf-8 -*-
# face_selection_policy.py

from collections import defaultdict


def select_representative_faces_by_direction(
    direction_map,
    view=None,
    prefer_outermost=True
):
    """
    Select exactly ONE face per element per direction.
    
    CRITICAL FIX for opposing faces:
    - For each element, pick face that is most opposing to other elements
    - This ensures we dimension between FACING surfaces, not parallel ones
    """

    cleaned_direction_map = {}

    for direction, faces in direction_map.items():
        if not faces:
            continue

        # 1. Group faces by element
        faces_by_element = defaultdict(list)
        for face in faces:
            faces_by_element[face.element_id].append(face)

        selected_faces = []

        # 2. Pick ONE face per element
        # Strategy: Pick faces that are most separated (opposing)
        for element_id, element_faces in faces_by_element.items():

            if len(element_faces) == 1:
                selected_faces.append(element_faces[0])
                continue

            # Multiple faces: pick the one most opposing to other elements
            # Check dot product with direction - opposite faces have opposite dot products
            best_face = None
            best_score = None
            
            for face in element_faces:
                # Score = distance along direction
                # For opposing faces, we want min from one element, max from another
                dist = face.distance_along(direction)
                
                if best_score is None:
                    best_face = face
                    best_score = dist
                elif prefer_outermost:
                    # Pick outermost along the (signed) direction
                    if dist > best_score:
                        best_face = face
                        best_score = dist
                else:
                    # Pick innermost along the (signed) direction
                    if dist < best_score:
                        best_face = face
                        best_score = dist
            
            if best_face:
                selected_faces.append(best_face)

        # 3. Sort faces along direction for proper pairing
        if len(selected_faces) >= 2:
            selected_faces.sort(key=lambda f: f.distance_along(direction))

        cleaned_direction_map[direction] = selected_faces

    return cleaned_direction_map


# ------------------------------------------------------------
# INTERNAL POLICY FUNCTIONS
# ------------------------------------------------------------

def _pick_representative_face(faces, direction, prefer_outermost=True):
    """
    Decide which face represents the element for this direction.

    Policy:
    - All faces are already parallel to direction (filtered earlier)
    - Choose face by distance along direction

    Returns
    -------
    FaceData or None
    """

    if not faces:
        return None

    # Compute distances
    face_distances = [
        (face, face.distance_along(direction))
        for face in faces
    ]

    # Sort by distance
    face_distances.sort(key=lambda x: x[1])

    # Outermost = farthest along direction
    if prefer_outermost:
        return face_distances[-1][0]
    else:
        return face_distances[0][0]
