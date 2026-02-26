# -*- coding: utf-8 -*-

from pyrevit import revit, forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import System
import traceback
import os

doc   = revit.doc
uidoc = revit.uidoc
app   = doc.Application


# ==========================================================
# SELECTION FILTER (ALLOW HOST + LINK)
# ==========================================================
class LinkAndHostFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return True

    def AllowReference(self, reference, position):
        return True


# ==========================================================
# PICK ELEMENTS
# ==========================================================
try:
    refs = uidoc.Selection.PickObjects(
        ObjectType.LinkedElement,
        LinkAndHostFilter(),
        "Pick elements (Host + Linked)"
    )
except:
    forms.alert("No element selected", exitscript=True)

if not refs or len(refs) == 0:
    forms.alert("No element selected", exitscript=True)


# ==========================================================
# SPLIT HOST / LINK ELEMENTS
# ==========================================================
host_elements   = []
linked_elements = []

for ref in refs:
    if ref.LinkedElementId != ElementId.InvalidElementId:
        link_inst = doc.GetElement(ref.ElementId)  # RevitLinkInstance
        link_doc  = link_inst.GetLinkDocument()
        linked_el = link_doc.GetElement(ref.LinkedElementId)
        linked_elements.append((linked_el, link_inst))
    else:
        host_elements.append(doc.GetElement(ref.ElementId))


# ==========================================================
# GET SOLIDS FROM ELEMENT
# ==========================================================
def get_solids(elem):
    solids = []
    opt = Options()
    opt.DetailLevel = ViewDetailLevel.Fine
    opt.IncludeNonVisibleObjects = True

    geo = elem.get_Geometry(opt)
    if not geo:
        return solids

    for g in geo:
        if isinstance(g, Solid) and g.Volume > 0:
            solids.append(g)
        elif isinstance(g, GeometryInstance):
            for ig in g.GetInstanceGeometry():
                if isinstance(ig, Solid) and ig.Volume > 0:
                    solids.append(ig)
    return solids


# ==========================================================
# COLLECT & TRANSFORM SOLIDS
# ==========================================================
all_solids = []

# Host solids
for el in host_elements:
    all_solids.extend(get_solids(el))

# Linked solids (apply transform)
for el, link_inst in linked_elements:
    tf = link_inst.GetTotalTransform()
    for s in get_solids(el):
        all_solids.append(SolidUtils.CreateTransformed(s, tf))


if not all_solids:
    forms.alert("No solid found in selected elements", exitscript=True)


# ==========================================================
# UNION SOLIDS
# ==========================================================
try:
    union_solid = all_solids[0]
    for s in all_solids[1:]:
        union_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
            union_solid, s, BooleanOperationsType.Union
        )
except:
    forms.alert(
        "Solid Union failed:\n{}".format(traceback.format_exc()),
        exitscript=True
    )


# ==========================================================
# INPUT FAMILY NAME WITH DUPLICATE CHECK
# ==========================================================
def get_existing_families():
    """Lấy danh sách tên families đã có trong project"""
    return [f.Name for f in FilteredElementCollector(doc).OfClass(Family)]

existing_families = get_existing_families()
default_name = "Void_From_SolidUnion"

# Tạo tên unique nếu trùng
fam_name = default_name
counter = 1
while fam_name in existing_families:
    fam_name = "{}_{}".format(default_name, counter)
    counter += 1

# Hiển thị UI để người dùng nhập tên
fam_name_input = forms.ask_for_string(
    default=fam_name,
    prompt="Enter Family Name:",
    title="Void Family Name"
)

if not fam_name_input:
    forms.alert("No name provided", exitscript=True)

fam_name = fam_name_input.strip()

# Kiểm tra tên trùng
if fam_name in existing_families:
    overwrite = forms.alert(
        "Family '{}' already exists!\n\nDo you want to overwrite?".format(fam_name),
        yes=True,
        no=True
    )
    if not overwrite:
        forms.alert("Cancelled", exitscript=True)


# ==========================================================
# CREATE VOID FAMILY
# ==========================================================
# ⚠️ FACE-BASED TEMPLATE (BẮT BUỘC ĐỂ CUT)
script_dir = os.path.dirname(__file__)
# Sử dụng Metric Generic Model face based template để tạo void cut
fam_template = os.path.join(script_dir, "Metric Generic Model face based.rft")

if not os.path.exists(fam_template):
    forms.alert(
        "Template not found:\n" + fam_template + 
        "\n\nPlease use 'Metric Generic Model face based.rft'",
        exitscript=True
    )

famdoc = app.NewFamilyDocument(fam_template)

t = Transaction(famdoc, "Create Void Solid")
t.Start()

# Create FreeForm from solid
ff = FreeFormElement.Create(famdoc, union_solid)

# Set as VOID
ff.get_Parameter(BuiltInParameter.ELEMENT_IS_CUTTING).Set(1)

# Allow family to cut
famdoc.OwnerFamily.get_Parameter(
    BuiltInParameter.FAMILY_ALLOW_CUT_WITH_VOIDS
).Set(1)

# Đổi tên family (QUAN TRỌNG!)
famdoc.OwnerFamily.Name = fam_name

t.Commit()


# ==========================================================
# SAVE FAMILY
# ==========================================================
save_path = os.path.join(script_dir, fam_name + ".rfa")

opt = SaveAsOptions()
opt.OverwriteExistingFile = True
famdoc.SaveAs(save_path, opt)
famdoc.Close(False)


# ==========================================================
# LOAD FAMILY (FIX LOADFAMILY BUG)
# ==========================================================
class FamOpt(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        return True
    def OnSharedFamilyFound(self, familyInUse, source, overwriteParameterValues):
        return True


t_load = Transaction(doc, "Load Void Family")
t_load.Start()
loaded_fam = doc.LoadFamily(save_path, FamOpt())
t_load.Commit()


# ==========================================================
# FIND LOADED FAMILY
# ==========================================================
fam = None
for f in FilteredElementCollector(doc).OfClass(Family):
    if f.Name == fam_name:
        fam = f
        break

# Nếu không tìm thấy, liệt kê tất cả families để debug
if not fam:
    all_fam_names = [f.Name for f in FilteredElementCollector(doc).OfClass(Family)]
    forms.alert(
        "Family not found after loading!\n\n"
        "Looking for: {}\n\n".format(fam_name) +
        "Available families:\n" + "\n".join(all_fam_names[-10:]),
        exitscript=True
    )

# Xóa file tạm sau khi đã tìm thấy family
try:
    System.IO.File.Delete(save_path)
except:
    pass  # Không quan trọng nếu xóa thất bại


# ==========================================================
# PLACE VOID INSTANCE
# ==========================================================
symbol_id = list(fam.GetFamilySymbolIds())[0]
symbol = doc.GetElement(symbol_id)

# Activate + Place trong cùng 1 transaction với error handling
t_place = Transaction(doc, "Activate & Place Void")
t_place.Start()

try:
    if not symbol.IsActive:
        symbol.Activate()
    
    doc.Create.NewFamilyInstance(
        XYZ.Zero,
        symbol,
        StructuralType.NonStructural
    )
    t_place.Commit()
except Exception as e:
    t_place.RollBack()
    forms.alert("Failed to place instance:\n{}".format(str(e)), exitscript=True)


forms.alert(
    "✅ VOID FAMILY CREATED\n\n"
    "Family: {}\n\n".format(fam_name) +
    "⚠️ IMPORTANT:\n"
    "This is a FACE-BASED family.\n"
    "To cut geometry:\n"
    "1. Select the void instance\n"
    "2. Place it on a wall/floor face\n"
    "3. Use 'Cut Geometry' tool"
)
# ==========================================================