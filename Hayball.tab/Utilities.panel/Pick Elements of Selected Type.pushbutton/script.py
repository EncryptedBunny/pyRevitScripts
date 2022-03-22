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
__title__ = 'Pick this\nElement'
__doc__ = 'Allows Selection of only this catergory'
__author__ = 'Sebastian Teo / pyRevit.Tools'
__helpurl__ = 'mailto:steo@hayball.com.au'

# Get Current Revit Document
doc = revit.doc

# Inputs
l = ['Area',
     'Area Boundary',
     'Column',
     'Dimension',
     'Door',
     'Floor',
     'Framing',
     'Furniture',
     'Grid',
     'Rooms',
     'Room Tag',
     'Truss',
     'Wall',
     'Window',
     'Ceiling',
     'Section Box',
     'Elevation Mark',
     'Parking']

l.sort()

class DetailSelection(UI.Selection.ISelectionFilter):
    # standard API override function
    def AllowElement(self, element):
        # only allow view-specific (detail) elements
        # that are not part of a group
        if element.ViewSpecific:
            if element.GroupId == element.GroupId.InvalidElementId:
                return True
        return False

    # standard API override function
    def AllowReference(self, refer, point):
        return False

class ModelSelection(UI.Selection.ISelectionFilter):
    # standard API override function
    def AllowElement(self, element):
        if not element.ViewSpecific:
            return True
        else:
            return False

    # standard API override function
    def AllowReference(self, refer, point):
        return False

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
        lab = Label(Text = 'Pick only elements of type:')
        self.Controls.Add(lab)
        lab.Size = Size(200,20)
        lab.Location = Point(10,12)
        self.lb.Parent = self
        self.lb.Size = Size(200,180)
        self.lb.Location = Point(10,35)

        self.ControlBox = True

        for data in dataEnteringNode:
            self.lb.Items.Add("%s" % data)
     
        done = Button(Text='Pick the Selected Catergory')
        self.Controls.Add(done)
        done.Location = Point(235, 15)
        done.Size = Size(75,50)
        done.Click += self.doneClick

        detail = Button(Text='Pick all Detail Elements')
        self.Controls.Add(detail)
        detail.Location = Point(235, 72)
        detail.Size = Size(75,50)
        detail.Click += self.detailClick

        model = Button(Text='Pick all Model Elements')
        self.Controls.Add(model)
        model.Location = Point(235, 129)
        model.Size = Size(75,50)
        model.Click += self.modelClick
        
        self.CenterToScreen()
            
    def doneClick(self, sender, event):
        
        global returnList
        returnList = "%s" % self.lb.SelectedItem
        self.Close()

    def detailClick(self, sender, event):
        global detail
        detail = True
        self.Close()
    
    def modelClick(self, sender, event):
        global model
        model = True
        self.Close()

# Run Checkbox with List
IForm(l).ShowDialog()

if 'detail' in globals():
    if detail:
        
        selection = revit.get_selection()

        try:
            msfilter = DetailSelection()
            selection_list = revit.pick_rectangle(pick_filter=msfilter)

            filtered_list = []
            for el in selection_list:
                filtered_list.append(el.Id)

            selection.set_to(filtered_list)
            revit.uidoc.RefreshActiveView()
        except Exception:
            pass

        sys.exit()
else:
    None

if 'model' in globals():
    if model:
        
        selection = revit.get_selection()
        
        try:
            msfilter = ModelSelection()
            selection_list = revit.pick_rectangle(pick_filter=msfilter)

            filtered_list = []
            for el in selection_list:
                filtered_list.append(el.Id)

            selection.set_to(filtered_list)
            revit.uidoc.RefreshActiveView()
        except Exception:
            pass

        sys.exit()
else:
    None

try:
    if returnList != "None":
        catSelect = returnList
    else:
        sys.exit()
except:
    sys.exit()

logger = script.get_logger()

class PickByCategorySelectionFilter(UI.Selection.ISelectionFilter):
    def __init__(self, catname):
        self.category = catname

    # standard API override function
    def AllowElement(self, element):
        if self.category in element.Category.Name:
            return True
        else:
            return False

    # standard API override function
    def AllowReference(self, refer, point): #pylint: disable=W0613
        return False

def pickbycategory(catname):
    try:
        selection = revit.get_selection()
        msfilter = PickByCategorySelectionFilter(catname)
        selection_list = revit.pick_rectangle(pick_filter=msfilter)
        filtered_list = []
        for element in selection_list:
            filtered_list.append(element.Id)
        selection.set_to(filtered_list)
    except Exception as err:
        logger.debug(err)

pickbycategory(catSelect)
