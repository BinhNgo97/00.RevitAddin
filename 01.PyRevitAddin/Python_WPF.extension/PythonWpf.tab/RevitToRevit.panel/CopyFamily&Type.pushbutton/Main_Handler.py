# -*- coding: utf-8 -*-
# Duplicate-name resolution used during CopyElements

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import IDuplicateTypeNamesHandler, DuplicateTypeAction

class DuplicatesHandler(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        # Keep destination definitions; skip incoming duplicates
        return DuplicateTypeAction.UseDestination