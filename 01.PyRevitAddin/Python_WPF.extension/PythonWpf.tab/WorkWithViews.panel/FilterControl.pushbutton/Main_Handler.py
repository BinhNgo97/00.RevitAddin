# -*- coding: utf-8 -*-
# Filter Handler - Using ISOLATE/UNISOLATE - Revit 2021 Compatible

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.UI import IExternalEventHandler
from Autodesk.Revit.DB import (
    Transaction,
    OverrideGraphicSettings,
    Color,
    FillPatternElement,
    FillPatternTarget,
    FilteredElementCollector,
)


class ExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        # inputs
        self.View = None
        self.FilterItems = []  # legacy: list of FilterItem for 'enable'
        self.FilterIds = []    # list[ElementId] for add/remove
        self.VisibilityChanges = []  # list[(ElementId, bool)] for apply_visibility
        self.ColorOverrides = []  # list[(ElementId, Color)] for set_filter_colors
        self.action = None  # 'enable' | 'apply_visibility' | 'add_filters' | 'remove_filters' | 'set_filter_colors'
        # outputs
        self.message = ""
        self.last_result = None  # True on success, False on error
        self.OnCompleted = None  # optional callback set by UI
        # busy flag
        self.is_busy = False
    
    def GetName(self):
        return "Filter Control Handler"
    
    def Execute(self, app):
        self.is_busy = True
        success = False
        try:
            if self.action == "enable":
                success = self._update_filters_enabled(app)
            elif self.action == "apply_visibility":
                success = self._apply_visibility(app)
            elif self.action == "add_filters":
                success = self._add_filters(app)
            elif self.action == "remove_filters":
                success = self._remove_filters(app)
            elif self.action == "set_filter_colors":
                success = self._set_filter_colors(app)
        except Exception as ex:
            self.message = "Error: {}".format(str(ex))
            print("Execute error: {}".format(str(ex)))
            success = False
        finally:
            self.last_result = success
            finished_action = self.action
            self.action = None
            # clear inputs after run
            self.FilterItems = []
            self.FilterIds = []
            self.VisibilityChanges = []
            self.ColorOverrides = []
            self.is_busy = False
            try:
                if self.OnCompleted is not None:
                    # Notify UI; pass success and action name
                    self.OnCompleted(success, finished_action)
            except Exception as cb_ex:
                print("OnCompleted callback error: {}".format(cb_ex))
    
    def _update_filters_enabled(self, app):
        """Enable/Disable multiple filters in one transaction - Revit 2021"""
        t = None
        try:
            if not self.View or not self.FilterItems:
                return True

            doc = app.ActiveUIDocument.Document

            t = Transaction(doc, "Toggle Filters Enabled")
            t.Start()

            # Apply visibility for each provided item
            for item in self.FilterItems:
                try:
                    self.View.SetFilterVisibility(item.FilterId, item.IsEnabled)
                except Exception as inner_ex:
                    print("Failed to set visibility for {}: {}".format(getattr(item, 'Name', item.FilterId), str(inner_ex)))

            t.Commit()
            return True

        except Exception as ex:
            if t and t.HasStarted():
                t.RollBack()
            print("Error updating filters enabled: {}".format(str(ex)))
            return False

    def _apply_visibility(self, app):
        t = None
        try:
            if not self.View or not self.VisibilityChanges:
                return True
            doc = app.ActiveUIDocument.Document
            t = Transaction(doc, "Apply Filter Visibility")
            t.Start()
            for (fid, vis) in self.VisibilityChanges:
                try:
                    self.View.SetFilterVisibility(fid, bool(vis))
                except Exception as inner_ex:
                    print("Failed SetFilterVisibility for {}: {}".format(fid, inner_ex))
            t.Commit()
            return True
        except Exception as ex:
            if t and t.HasStarted():
                t.RollBack()
            print("Error applying visibility: {}".format(str(ex)))
            return False

    def _add_filters(self, app):
        t = None
        try:
            if not self.View or not self.FilterIds:
                return True
            doc = app.ActiveUIDocument.Document
            t = Transaction(doc, "Add Filters To View")
            t.Start()
            for fid in self.FilterIds:
                try:
                    self.View.AddFilter(fid)
                except Exception as inner_ex:
                    print("Failed AddFilter for {}: {}".format(fid, inner_ex))
            t.Commit()
            return True
        except Exception as ex:
            if t and t.HasStarted():
                t.RollBack()
            print("Error adding filters: {}".format(str(ex)))
            return False

    def _remove_filters(self, app):
        t = None
        try:
            if not self.View or not self.FilterIds:
                return True
            doc = app.ActiveUIDocument.Document
            t = Transaction(doc, "Remove Filters From View")
            t.Start()
            for fid in self.FilterIds:
                try:
                    self.View.RemoveFilter(fid)
                except Exception as inner_ex:
                    print("Failed RemoveFilter for {}: {}".format(fid, inner_ex))
            t.Commit()
            return True
        except Exception as ex:
            if t and t.HasStarted():
                t.RollBack()
            print("Error removing filters: {}".format(str(ex)))
            return False

    def _set_filter_colors(self, app):
        """Apply Solid Fill foreground pattern color (Projection/Surface),
        with best-effort compatibility for different Revit versions.
        """
        t = None
        try:
            if not self.View or not self.ColorOverrides:
                return True

            doc = app.ActiveUIDocument.Document

            # Resolve a Solid Fill pattern element (prefer Drafting target)
            solid_fill_id = None
            try:
                for fpe in FilteredElementCollector(doc).OfClass(FillPatternElement):
                    try:
                        fp = fpe.GetFillPattern()
                        if fp and getattr(fp, 'IsSolidFill', False):
                            # Prefer Drafting target
                            if getattr(fp, 'Target', FillPatternTarget.Drafting) == FillPatternTarget.Drafting:
                                solid_fill_id = fpe.Id
                                break
                            if solid_fill_id is None:
                                solid_fill_id = fpe.Id
                    except Exception:
                        pass
            except Exception as find_ex:
                print("Could not enumerate FillPatternElement: {}".format(find_ex))

            t = Transaction(doc, "Set Filter Pattern Colors")
            t.Start()

            for (fid, col) in self.ColorOverrides:
                try:
                    ogs = self.View.GetFilterOverrides(fid) or OverrideGraphicSettings()

                    # Try newer Surface Foreground API first
                    applied = False
                    try:
                        if solid_fill_id:
                            ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                            ogs.SetSurfaceForegroundPatternVisible(True)
                        ogs.SetSurfaceForegroundPatternColor(col)
                        applied = True
                    except Exception:
                        # Fallback to Projection Fill API (older)
                        try:
                            if solid_fill_id:
                                ogs.SetProjectionFillPatternId(solid_fill_id)
                                ogs.SetProjectionFillPatternVisible(True)
                            ogs.SetProjectionFillPatternColor(col)
                            applied = True
                        except Exception as cex:
                            print("Set projection/surface pattern color failed for {}: {}".format(fid, cex))

                    # Optionally set Cut foreground pattern color if available
                    try:
                        if solid_fill_id:
                            ogs.SetCutForegroundPatternId(solid_fill_id)
                            ogs.SetCutForegroundPatternVisible(True)
                            ogs.SetCutForegroundPatternColor(col)
                    except Exception:
                        # Fallback older cut API if present
                        try:
                            ogs.SetCutFillPatternColor(col)
                        except Exception:
                            pass

                    self.View.SetFilterOverrides(fid, ogs)
                except Exception as inner_ex:
                    print("Failed SetFilterOverrides for {}: {}".format(fid, inner_ex))

            t.Commit()
            return True

        except Exception as ex:
            if t and t.HasStarted():
                t.RollBack()
            print("Error setting filter colors: {}".format(str(ex)))
            return False
