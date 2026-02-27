# -*- coding: utf-8 -*-
import sys
import os
import clr
clr.AddReference('System')

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from pyrevit import forms
from ViewModel import LinkMultiCadViewModel
import Handler

try:
    from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
    from Autodesk.Revit.DB import FilteredElementCollector, Level
    _REVIT_AVAILABLE = True
except Exception:
    _REVIT_AVAILABLE = False


def _load_revit_levels(vm):
    if not _REVIT_AVAILABLE:
        return
    try:
        doc = __revit__.ActiveUIDocument.Document  # noqa: F821

        level_elems = list(
            FilteredElementCollector(doc).OfClass(Level).ToElements()
        )

        def _elev(lv):
            try:
                return lv.Elevation
            except Exception:
                return 0.0

        names = []
        for lv in sorted(level_elems, key=_elev):
            try:
                names.append(lv.Name)
            except Exception:
                pass

        vm.set_revit_levels(names)

    except Exception as ex:
        print("_load_revit_levels error: {}".format(ex))


class LinkMultiCadWindow(forms.WPFWindow):

    def __init__(self, xaml_file, view_model):
        try:
            forms.WPFWindow.__init__(self, xaml_file)
            self.DataContext = view_model

            if _REVIT_AVAILABLE:
                select_baseline_handler = Handler.SelectBaseLineHandler(view_model)
                link_cad_handler        = Handler.LinkCadHandler(view_model)

                self._ext_select_baseline = ExternalEvent.Create(select_baseline_handler)
                self._ext_link_cad        = ExternalEvent.Create(link_cad_handler)

                _load_revit_levels(view_model)

            Handler.bind_handlers(self)

            view_model.Status = u"Ready."

        except Exception as ex:
            print("LinkMultiCadWindow init error: {}".format(ex))
            raise


if __name__ == '__main__':
    vm     = LinkMultiCadViewModel()
    window = LinkMultiCadWindow('UI.xaml', vm)
    window.Show()
