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
__title__ = 'Create Path\nfor Labels'
__doc__ = "Creates Line Style 'Label Maker', which can be used to draw the determined path on which labels are automatically assigned and incremented."
__author__ = 'Sebastian Teo'
__helpurl__ = 'mailto:steo@hayball.com.au'

# Get Current Revit Document
doc = revit.doc
uiapp = HOST_APP.uiapp
# Line Style
patterns = DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement)
for i in patterns:
    if i.Name == 'Long dash':
        pattern = i

categories = doc.Settings.Categories
lineCat = categories.get_Item(DB.BuiltInCategory.OST_Lines)

with DB.Transaction(doc, "Create Linestyle") as trans:
    trans.Start()

    try:
        newLineStyleCat = categories.NewSubcategory(lineCat, "Label Maker")
    except:
        for i in lineCat.SubCategories:
            if i.Name == "Label Maker":
                newLineStyleCat = i

    doc.Regenerate()

    newLineStyleCat.SetLineWeight(10,DB.GraphicsStyleType.Projection)
    newLineStyleCat.LineColor = DB.Color(0xFF, 0x00, 0xFF)
    newLineStyleCat.SetLinePatternId(pattern.Id, DB.GraphicsStyleType.Projection)

    trans.Commit()

# Post

class IForm(Form):
        
    def __init__(self):

        # Logo
        logo = Panel()
        self.Controls.Add(logo)
        img = Image.FromFile("%s" % os.path.dirname(os.path.realpath(__file__)) + "\\Hayball_Logo.png")
        logo.BackgroundImage = img
        logo.Location = Point(10, 180)
        logo.Size = Size(200,50)

        pnl = Panel()
        self.Controls.Add(pnl)
        img = Image.FromFile("%s" % os.path.dirname(os.path.realpath(__file__)) + "\\LineStyle.png")
        pnl.BackgroundImage = img
        pnl.Location = Point(10, 68)
        pnl.Size = Size(290,90)
        
        self.Size = Size(330, 280)
        self.Text = "Warning"

        lb = Label(Text = "Draw the path of order to Label.\n\nPlease change 'Line Style' to 'Label Maker' before continuing")
        self.Controls.Add(lb)
        lb.Size = Size(200,70)
        lb.Location = Point(15,8)
        
        self.ControlBox = False

        done = Button(Text='Got it!')
        self.Controls.Add(done)
        done.Location = Point(215, 20)
        done.Click += self.doneClick
        
        self.CenterToScreen()

    def radioClick(self, sender, event):
        rb = sender
        if rb.Checked:
            self.selected = rb
            
    def doneClick(self, sender, event):
        
        self.Close()

IForm().ShowDialog()

# Data Collection

script.get_results().Calculations = 0
script.get_results().Minutes = 0


detailCmd = UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.DetailLine)
uiapp.PostCommand(detailCmd)
