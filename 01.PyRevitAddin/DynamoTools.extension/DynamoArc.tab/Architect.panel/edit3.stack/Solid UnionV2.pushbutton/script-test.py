# -*- coding: utf-8 -*-

from pyrevit import revit, forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import System
from System.Windows import Window, Application
from System.Windows.Controls import Button, TextBox, TextBlock, StackPanel
from System.Windows.Media import SolidColorBrush, Color
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment, FontWeight
import traceback
import os

doc = revit.doc
uidoc = revit.uidoc
app = doc.Application


# ==========================================================
# SELECTION FILTER
# ==========================================================
class LinkAndHostFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return True

    def AllowReference(self, reference, position):
        return True


# ==========================================================
# MODELESS WPF WINDOW
# ==========================================================
class SolidUnionWindow(Window):
    def __init__(self):
        self.Title = "Solid Union - Void Family Creator"
        self.Width = 500
        self.Height = 400
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Color.FromRgb(240, 240, 240))
        
        # Data storage
        self.selected_refs = None
        
        # Create UI
        self.create_ui()
    
    def create_ui(self):
        # Main StackPanel
        main_panel = StackPanel()
        main_panel.Margin = Thickness(20)
        
        # === SELECT BUTTON ===
        self.btn_select = Button()
        self.btn_select.Content = "Select link element"
        self.btn_select.Height = 50
        self.btn_select.FontSize = 16
        self.btn_select.FontWeight = FontWeight.FromOpenTypeWeight(600)
        self.btn_select.Background = SolidColorBrush(Color.FromRgb(0, 191, 255))
        self.btn_select.Foreground = SolidColorBrush(Color.FromRgb(0, 0, 0))
        self.btn_select.Margin = Thickness(0, 0, 0, 10)
        self.btn_select.Click += self.on_select_click
        main_panel.Children.Add(self.btn_select)
        
        # === FAMILY NAME TEXTBOX ===
        self.txt_family_name = TextBox()
        self.txt_family_name.Text = "Void_From_SolidUnion"
        self.txt_family_name.Height = 50
        self.txt_family_name.FontSize = 16
        self.txt_family_name.Background = SolidColorBrush(Color.FromRgb(0, 191, 255))
        self.txt_family_name.Foreground = SolidColorBrush(Color.FromRgb(0, 0, 0))
        self.txt_family_name.Padding = Thickness(10)
        self.txt_family_name.Margin = Thickness(0, 0, 0, 10)
        self.txt_family_name.VerticalContentAlignment = VerticalAlignment.Center
        main_panel.Children.Add(self.txt_family_name)
        
        # === CREATE BUTTON ===
        self.btn_create = Button()
        self.btn_create.Content = "Create"
        self.btn_create.Height = 50
        self.btn_create.FontSize = 16
        self.btn_create.FontWeight = FontWeight.FromOpenTypeWeight(600)
        self.btn_create.Background = SolidColorBrush(Color.FromRgb(0, 191, 255))
        self.btn_create.Foreground = SolidColorBrush(Color.FromRgb(0, 0, 0))
        self.btn_create.Margin = Thickness(0, 0, 0, 20)
        self.btn_create.Click += self.on_create_click
        main_panel.Children.Add(self.btn_create)
        
        # === RESULT TEXTBLOCK ===
        self.txt_result = TextBlock()
        self.txt_result.Text = "Family was created:\nName:"
        self.txt_result.FontSize = 14
        self.txt_result.Background = SolidColorBrush(Color.FromRgb(0, 191, 255))
        self.txt_result.Foreground = SolidColorBrush(Color.FromRgb(0, 0, 0))
        self.txt_result.Padding = Thickness(15)
        self.txt_result.TextWrapping = System.Windows.TextWrapping.Wrap
        self.txt_result.MinHeight = 100
        main_panel.Children.Add(self.txt_result)
        
        self.Content = main_panel
    
    def on_select_click(self, sender, args):
        """Xử lý khi click Select button"""
        try:
            # Hide window temporarily
            self.Hide()
            
            # Pick elements
            self.selected_refs = uidoc.Selection.PickObjects(
                ObjectType.LinkedElement,
                LinkAndHostFilter(),
                "Pick elements (Host + Linked)"
            )
            
            if self.selected_refs and len(self.selected_refs) > 0:
                self.txt_result.Text = "Selected {} element(s)\nReady to create!".format(
                    len(self.selected_refs)
                )
            else:
                self.txt_result.Text = "No elements selected"
            
        except Exception as e:
            self.txt_result.Text = "Selection cancelled or error:\n{}".format(str(e))
        finally:
            # Show window again
            self.Show()
    
    def on_create_click(self, sender, args):
        """Xử lý khi click Create button"""
        if not self.selected_refs or len(self.selected_refs) == 0:
            self.txt_result.Text = "ERROR: Please select elements first!"
            return
        
        fam_name = self.txt_family_name.Text.strip()
        if not fam_name:
            self.txt_result.Text = "ERROR: Please enter family name!"
            return
        
        try:
            # Process and create family
            result = self.create_void_family(self.selected_refs, fam_name)
            
            if result:
                self.txt_result.Text = "Family was created:\nName: {}".format(fam_name)
            else:
                self.txt_result.Text = "ERROR: Failed to create family"
                
        except Exception as e:
            self.txt_result.Text = "ERROR:\n{}".format(str(e))
    
    def create_void_family(self, refs, fam_name):
        """Tạo void family từ selected elements"""
        # Split host/link elements
        host_elements = []
        linked_elements = []
        
        for ref in refs:
            if ref.LinkedElementId != ElementId.InvalidElementId:
                link_inst = doc.GetElement(ref.ElementId)
                link_doc = link_inst.GetLinkDocument()
                linked_el = link_doc.GetElement(ref.LinkedElementId)
                linked_elements.append((linked_el, link_inst))
            else:
                host_elements.append(doc.GetElement(ref.ElementId))
        
        # Get solids
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
        
        # Collect & transform solids
        all_solids = []
        
        for el in host_elements:
            all_solids.extend(get_solids(el))
        
        for el, link_inst in linked_elements:
            tf = link_inst.GetTotalTransform()
            for s in get_solids(el):
                all_solids.append(SolidUtils.CreateTransformed(s, tf))
        
        if not all_solids:
            raise Exception("No solid found in selected elements")
        
        # Union solids
        union_solid = all_solids[0]
        for s in all_solids[1:]:
            union_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                union_solid, s, BooleanOperationsType.Union
            )
        
        # Create family
        script_dir = os.path.dirname(__file__)
        fam_template = os.path.join(script_dir, "Metric Generic Model face based.rft")
        
        if not os.path.exists(fam_template):
            raise Exception("Template not found:\n" + fam_template)
        
        famdoc = app.NewFamilyDocument(fam_template)
        
        t = Transaction(famdoc, "Create Void Solid")
        t.Start()
        
        ff = FreeFormElement.Create(famdoc, union_solid)
        ff.get_Parameter(BuiltInParameter.ELEMENT_IS_CUTTING).Set(1)
        famdoc.OwnerFamily.get_Parameter(
            BuiltInParameter.FAMILY_ALLOW_CUT_WITH_VOIDS
        ).Set(1)
        famdoc.OwnerFamily.Name = fam_name
        
        t.Commit()
        
        # Save family
        save_path = os.path.join(script_dir, fam_name + ".rfa")
        opt = SaveAsOptions()
        opt.OverwriteExistingFile = True
        famdoc.SaveAs(save_path, opt)
        famdoc.Close(False)
        
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
            raise Exception("Family not found after loading")
        
        # Delete temp file
        try:
            System.IO.File.Delete(save_path)
        except:
            pass
        
        # Place instance
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
        except Exception as e:
            t_place.RollBack()
            raise e
        
        return True


# ==========================================================
# MAIN - SHOW MODELESS WINDOW
# ==========================================================
window = SolidUnionWindow()
window.Show()
