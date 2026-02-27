# -*- coding: utf-8 -*-
"""
EXAMPLE: How to implement View Change Detection
================================================

This file shows optional implementation for automatically refreshing 
lstProjectFilters and lstViewFilters when user switches views in Revit.

CURRENT BEHAVIOR:
-----------------
✓ Auto-refresh after add/remove filters (already implemented)
✗ Does NOT auto-refresh when user switches to different view

IMPLEMENTATION OPTIONS:
=======================

Option 1: MANUAL REFRESH BUTTON (Simplest)
-------------------------------------------
Add a button to XAML:
    <Button x:Name="btnRefresh" Content="🔄 Refresh" Click="OnRefreshClick"/>

Add handler in script.py:
    def OnRefreshClick(self, sender, e):
        self._load_collections()

Option 2: POLLING TIMER (Recommended)
--------------------------------------
Add to FilterControlWindow.__init__() after self._load_collections():
"""

# EXAMPLE CODE - Add to script.py FilterControlWindow.__init__():
"""
        from System import TimeSpan
        from System.Windows.Threading import DispatcherTimer

        # Track current view to detect changes
        self._last_view_id = None
        try:
            if revit.doc.ActiveView:
                self._last_view_id = revit.doc.ActiveView.Id.IntegerValue
        except:
            pass

        # Setup polling timer (checks every 1.5 seconds)
        self._view_check_timer = DispatcherTimer()
        self._view_check_timer.Interval = TimeSpan.FromSeconds(1.5)
        self._view_check_timer.Tick += self._check_view_changed
        self._view_check_timer.Start()

    def _check_view_changed(self, sender, e):
        '''Called by timer to detect view changes.'''
        try:
            current_view = None
            current_view_id = None
            try:
                current_view = revit.doc.ActiveView
                if current_view:
                    current_view_id = current_view.Id.IntegerValue
            except:
                pass

            # Compare with last known view
            if current_view_id != self._last_view_id:
                print("View changed detected - refreshing lists...")
                self._last_view_id = current_view_id
                
                # Update handler's view reference
                self._handler.View = current_view
                
                # Refresh both lists
                self._load_collections()
                
        except Exception as ex:
            print("View change check error: {}".format(str(ex)))
"""

# Option 3: REVIT VIEW ACTIVATION EVENT (Advanced, may not work in modeless)
# ----------------------------
# This would be ideal but requires access to UIApplication events
# which may not be available in pyRevit modeless windows

"""
SUMMARY OF REFRESH TRIGGERS:
=============================

Currently Implemented:
✓ Window opens → _load_collections()
✓ After add_filters → _on_handler_completed() → _load_collections()
✓ After remove_filters → _on_handler_completed() → _load_collections()
✓ After apply_visibility → _on_handler_completed() → _load_collections()
✓ After set_filter_colors → _on_handler_completed() → _load_collections()

Need to Add (choose one):
○ Manual refresh button (simple, reliable)
○ Polling timer (automatic, uses ~0.5% CPU)
○ View activation event (ideal but may not be available)

RECOMMENDATION:
===============
Use Option 2 (Polling Timer) for best user experience.
It's automatic, reliable, and has minimal performance impact.
"""
