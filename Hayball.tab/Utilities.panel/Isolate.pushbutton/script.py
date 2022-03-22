#pylint: disable=E0401,C0103,maybe-no-member

# Import pyRevit modules
from pyrevit import script
from pyrevit.framework import List
from pyrevit.framework import clr
from pyrevit import revit, DB, HOST_APP, UI

# Import Winform Modules
clr.AddReference("System.Windows.Forms")
clr.AddReference("system.Drawing")
clr.AddReference('IronPython')
from System.Windows.Forms import Panel, GroupBox, RadioButton, ProgressBar, Application, Form, CheckBox, CheckedListBox, ListBox, Button, Label, TextBox
from System.Drawing import PointF, Size, Point, Font, FontStyle, Brushes, Image
from System.Threading import ThreadStart, Thread
from IronPython.Compiler import CallTarget0

# Import System modules
import System
import sys
import os
from System.Collections.Generic import *

# Set component Metadata
__title__ = 'Isolate this\nElement'
__doc__ = 'Isolates specific elements in current view and put the view in isolate element mode.'
__author__ = 'Sebastian Teo / pyRevit.Tools'
__helpurl__ = 'mailto:steo@hayball.com.au'

element_cats = {'Area Lines': DB.BuiltInCategory.OST_AreaSchemeLines,
                'Doors': DB.BuiltInCategory.OST_Doors,
                'Room Separation Lines':
                    DB.BuiltInCategory.OST_RoomSeparationLines,
                'Room Tags': None,
                'Model Groups': None,
                'Painted Elements': None,
                'Model Elements': None}

# List Box Class
class IForm(Form):
    
    lb = ListBox()
    
    def __init__(self, dataEnteringNode):

        # Logo
        logo = Panel()
        self.Controls.Add(logo)
        img = Image.FromFile("%s" % os.path.dirname(os.path.realpath(__file__)) + "\\Hayball_Logo.png")
        logo.BackgroundImage = img
        logo.Location = Point(10, 220)
        logo.Size = Size(200,50)
        
        self.Text = "Select Options"
        self.Size = Size(350, 320)
        lab = Label(Text = 'Isolate elements of type:')
        self.Controls.Add(lab)
        lab.Size = Size(200,20)
        lab.Location = Point(10,12)
        self.lb.Parent = self
        self.lb.Size = Size(200,180)
        self.lb.Location = Point(10,35)

        self.ControlBox = True

        for data in dataEnteringNode:
            self.lb.Items.Add("%s" % data)
     
        done = Button(Text='Done')
        self.Controls.Add(done)
        done.Location = Point(235, 35)
        done.Click += self.doneClick
        
        self.CenterToScreen()
            
    def doneClick(self, sender, event):
        
        global returnList
        returnList = "%s" % self.lb.SelectedItem
        self.Close()


# Run Checkbox with List
IForm(sorted(element_cats.keys())).ShowDialog()

try:
    if returnList != "None":
        selected_switch = returnList
    else:
        sys.exit()
except:
    sys.exit()

if selected_switch:
    curview = revit.activeview

    if selected_switch == 'Room Tags':
        roomtags = DB.FilteredElementCollector(revit.doc, curview.Id)\
                     .OfCategory(DB.BuiltInCategory.OST_RoomTags)\
                     .WhereElementIsNotElementType()\
                     .ToElementIds()

        rooms = DB.FilteredElementCollector(revit.doc, curview.Id)\
                  .OfCategory(DB.BuiltInCategory.OST_Rooms)\
                  .WhereElementIsNotElementType()\
                  .ToElementIds()

        allelements = []
        allelements.extend(rooms)
        allelements.extend(roomtags)
        element_to_isolate = List[DB.ElementId](allelements)

    elif selected_switch == 'Model Groups':
        elements = DB.FilteredElementCollector(revit.doc, curview.Id)\
                     .WhereElementIsNotElementType()\
                     .ToElementIds()

        modelgroups = []
        expanded = []
        for elid in elements:
            el = revit.doc.GetElement(elid)
            if isinstance(el, DB.Group) and not el.ViewSpecific:
                modelgroups.append(elid)
                members = el.GetMemberIds()
                expanded.extend(list(members))

        expanded.extend(modelgroups)
        element_to_isolate = List[DB.ElementId](expanded)

    elif selected_switch == 'Painted Elements':
        set = []

        elements = DB.FilteredElementCollector(revit.doc, curview.Id)\
                     .WhereElementIsNotElementType()\
                     .ToElementIds()

        for elId in elements:
            el = revit.doc.GetElement(elId)
            if len(list(el.GetMaterialIds(True))) > 0:
                set.append(elId)
            elif isinstance(el, DB.Wall) and el.IsStackedWall:
                memberWalls = el.GetStackedWallMemberIds()
                for mwid in memberWalls:
                    mw = revit.doc.GetElement(mwid)
                    if len(list(mw.GetMaterialIds(True))) > 0:
                        set.append(elId)
        element_to_isolate = List[DB.ElementId](set)

    elif selected_switch == 'Model Elements':
        elements = DB.FilteredElementCollector(revit.doc, curview.Id)\
                     .WhereElementIsNotElementType()\
                     .ToElementIds()

        element_to_isolate = []
        for elid in elements:
            el = revit.doc.GetElement(elid)
            if not el.ViewSpecific:  # and not isinstance(el, DB.Dimension):
                element_to_isolate.append(elid)

        element_to_isolate = List[DB.ElementId](element_to_isolate)

    else:
        element_to_isolate = \
            DB.FilteredElementCollector(revit.doc, curview.Id)\
              .OfCategory(element_cats[selected_switch]) \
              .WhereElementIsNotElementType()\
              .ToElementIds()

    # now that list of elements is ready, let's isolate them in the active view
    with revit.Transaction('Isolate {}'.format(selected_switch)):
        curview.IsolateElementsTemporary(element_to_isolate)
