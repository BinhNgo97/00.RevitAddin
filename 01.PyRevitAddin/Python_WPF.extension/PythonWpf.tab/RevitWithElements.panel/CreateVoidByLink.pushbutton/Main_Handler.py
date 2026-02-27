# -*- coding: utf-8 -*-
# CreateVoidByLink Handler - ExternalEventHandler for modeless operation

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.UI import IExternalEventHandler
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import System
import os
import traceback


class LinkAndHostFilter(ISelectionFilter):
    """Selection filter that allows both host and linked elements"""
    def AllowElement(self, elem):
        return True

    def AllowReference(self, reference, position):
        return True


class ExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        # inputs
        self.action = None  # 'select_elements' | 'create_void'
        self.FamilyName = None
        self.SelectedElements = []  # List of tuples: (ElementId, LinkedElementId or None)
        self.ConfirmOverwrite = True  # Set to True if user confirmed overwrite
        
        # outputs
        self.message = ""
        self.last_result = None  # True on success, False on error
        self.CreatedFamilyName = None
        self.OnCompleted = None  # optional callback set by UI
        
        # busy flag
        self.is_busy = False
    
    def GetName(self):
        return "CreateVoidByLink Handler"
    
    def Execute(self, app):
        self.is_busy = True
        success = False
        try:
            if self.action == "select_elements":
                success = self._select_elements(app)
            elif self.action == "create_void":
                success = self._create_void(app)
        except Exception as ex:
            self.message = "Error: {}".format(str(ex))
            print("Execute error: {}".format(str(ex)))
            traceback.print_exc()
            success = False
        finally:
            self.last_result = success
            finished_action = self.action
            self.action = None
            self.is_busy = False
            try:
                if self.OnCompleted is not None:
                    self.OnCompleted(success, finished_action)
            except Exception as cb_ex:
                print("OnCompleted callback error: {}".format(cb_ex))
    
    def _select_elements(self, app):
        """Select host and/or linked elements"""
        try:
            uidoc = app.ActiveUIDocument
            
            refs = uidoc.Selection.PickObjects(
                ObjectType.LinkedElement,
                LinkAndHostFilter(),
                "Pick elements (Host + Linked)"
            )
            
            if refs and len(refs) > 0:
                # Store ElementIds instead of Reference objects
                self.SelectedElements = []
                for ref in refs:
                    elem_id = ref.ElementId
                    linked_elem_id = ref.LinkedElementId if ref.LinkedElementId != ElementId.InvalidElementId else None
                    self.SelectedElements.append((elem_id, linked_elem_id))
                
                self.message = "Selected {} element(s)".format(len(refs))
                return True
            else:
                self.message = "No element selected"
                return False
                
        except Exception as ex:
            self.message = "Selection cancelled or failed: {}".format(str(ex))
            return False
    
    def _create_void(self, app):
        """Create void family from selected elements"""
        t = None
        famdoc = None
        save_path = None
        
        try:
            doc = app.ActiveUIDocument.Document
            
            if not self.SelectedElements or len(self.SelectedElements) == 0:
                self.message = "No elements selected. Please select elements first."
                return False
            
            if not self.FamilyName or not self.FamilyName.strip():
                self.message = "Please enter a family name"
                return False
            
            fam_name = self.FamilyName.strip()
            
            # Split host / link elements
            host_elements = []
            linked_elements = []
            
            for elem_id, linked_elem_id in self.SelectedElements:
                if linked_elem_id is not None:
                    # Linked element
                    link_inst = doc.GetElement(elem_id)
                    if link_inst is None:
                        continue
                    link_doc = link_inst.GetLinkDocument()
                    if link_doc is None:
                        continue
                    linked_el = link_doc.GetElement(linked_elem_id)
                    if linked_el is not None:
                        linked_elements.append((linked_el, link_inst))
                else:
                    # Host element
                    host_el = doc.GetElement(elem_id)
                    if host_el is not None:
                        host_elements.append(host_el)
            
            # Get solids from elements
            all_solids = []
            
            # Host solids
            for el in host_elements:
                all_solids.extend(self._get_solids(el))
            
            # Linked solids (apply transform)
            for el, link_inst in linked_elements:
                tf = link_inst.GetTotalTransform()
                for s in self._get_solids(el):
                    all_solids.append(SolidUtils.CreateTransformed(s, tf))
            
            if not all_solids:
                self.message = "No solid found in selected elements"
                return False
            
            # Union solids
            try:
                union_solid = all_solids[0]
                for s in all_solids[1:]:
                    union_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                        union_solid, s, BooleanOperationsType.Union
                    )
            except Exception as ex:
                self.message = "Solid Union failed: {}".format(str(ex))
                return False
            
            # Check if family exists
            existing_families = [f.Name for f in FilteredElementCollector(doc).OfClass(Family)]
            if fam_name in existing_families:
                # Family exists - always overwrite since we can't ask user from here
                self.message = "Family '{}' already exists, overwriting...".format(fam_name)
            
            # Create void family
            script_dir = os.path.dirname(__file__)
            fam_template = os.path.join(script_dir, "Metric Generic Model face based.rft")
            
            if not os.path.exists(fam_template):
                self.message = "Template not found: {}".format(fam_template)
                return False
            
            famdoc = app.Application.NewFamilyDocument(fam_template)
            
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
            
            # Set family name
            famdoc.OwnerFamily.Name = fam_name
            
            t.Commit()
            
            # Save family
            save_path = os.path.join(script_dir, fam_name + ".rfa")
            
            opt = SaveAsOptions()
            opt.OverwriteExistingFile = True
            famdoc.SaveAs(save_path, opt)
            famdoc.Close(False)
            famdoc = None
            
            # Load family
            class FamOpt(IFamilyLoadOptions):
                def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                    return True
                def OnSharedFamilyFound(self, familyInUse, source, overwriteParameterValues):
                    return True
            
            t_load = Transaction(doc, "Load Void Family")
            t_load.Start()
            doc.LoadFamily(save_path, FamOpt())
            t_load.Commit()
            
            # Find loaded family
            fam = None
            for f in FilteredElementCollector(doc).OfClass(Family):
                if f.Name == fam_name:
                    fam = f
                    break
            
            if not fam:
                self.message = "Family not found after loading!"
                return False
            
            # Delete temporary file
            try:
                System.IO.File.Delete(save_path)
            except:
                pass
            
            # Place void instance
            symbol_id = list(fam.GetFamilySymbolIds())[0]
            symbol = doc.GetElement(symbol_id)
            
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
            except Exception as ex:
                t_place.RollBack()
                self.message = "Failed to place instance: {}".format(str(ex))
                return False
            
            self.CreatedFamilyName = fam_name
            self.message = "Void family '{}' created successfully!".format(fam_name)
            return True
            
        except Exception as ex:
            if t and t.HasStarted():
                t.RollBack()
            if famdoc:
                try:
                    famdoc.Close(False)
                except:
                    pass
            self.message = "Error creating void: {}".format(str(ex))
            traceback.print_exc()
            return False
    
    def _get_solids(self, elem):
        """Extract solids from element geometry"""
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
