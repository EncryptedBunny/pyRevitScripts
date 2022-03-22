#pylint: disable=E0401,C0103,maybe-no-member

# Import pyRevit modules
from pyrevit import script
from pyrevit.framework import List
from pyrevit.framework import clr
from pyrevit import revit, DB, HOST_APP

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
__title__ = 'Area\nItem No'
__doc__ = 'Moves either Model Group Name or Area Name to the Area Item No for tagging. Room-Apartment Number of Casework in Model group must match Room-Apartment Number of Area to transfer Model Group name.'
__author__ = 'Sebastian Teo'
__helpurl__ = 'mailto:steo@hayball.com.au'

# Get Current Revit Document
doc = revit.doc

# Define Area names to collect
filterList = ["1B", "2B", "3B", "TH", "STUDIO", "TWIN", "CLUSTER"]

casework = []
apartmentNo = []
areas = []

for a in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToElements():
    for filter in filterList:
        if filter in a.GetParameters("Name")[0].AsString():
            if a.GetParameters("Area")[0].AsDouble() != 0:
                apt = a.GetParameters("Room-Apartment Number")[0].AsString()
                if apt != "":
                    areas.append(a)
                    apartmentNo.append(apt)

for i in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Casework).WhereElementIsNotElementType().ToElements():
    if i.GetParameters("Room-Apartment Number")[0].AsString() != "":
         casework.append(i)

l =["Model Group 'Name'", "Area 'Name'"]

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
        lab = Label(Text = 'Select Parameter to transfer From')
        self.Controls.Add(lab)
        lab.Size = Size(200,20)
        lab.Location = Point(10,12)
        self.lb.Parent = self
        self.lb.Size = Size(200,120)
        self.lb.Location = Point(10,35)

        self.ControlBox = True

        for data in dataEnteringNode:
            self.lb.Items.Add("%s" % data)
     
        done = Button(Text='Done')
        self.Controls.Add(done)
        done.Location = Point(235, 15)
        done.Click += self.doneClick
        
        self.CenterToScreen()
            
    def doneClick(self, sender, event):
        
        global returnList
        returnList = "%s" % self.lb.SelectedItem
        self.Close()

# Run Checkbox with List
IForm(l).ShowDialog()

try:
    if returnList != "":
        parameterToChange = returnList
    else:
        sys.exit()
except:
    sys.exit()

# Check
user = HOST_APP.username
elementsToCheck = areas

for element in elementsToCheck:
    bufstr = clr.Reference[str]()
    DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id, bufstr)
    _bufstr = '%s' % bufstr
    if _bufstr != "''":
        if _bufstr != "'%s'" % user:
            print("%s owns an element that will be modified. Please request release of element before trying again" % _bufstr)
            sys.exit()

# Progress Bar Class
class ProgressBarDialog(Form):
    
    # Initialization
    def __init__(self, numb, txt):
        
        self.Text = txt
        pb = ProgressBar()
        pb.Minimum = 1
        pb.Maximum = numb
        pb.Step = 1
        pb.Value = 1   
        pb.Width = 400
        self.ControlBox = False

        self.Controls.Add(pb)
        
        pb.Location = Point(5, 15)
        self.prog = pb
        self.Height = 100
        self.Width = 425
        self.CenterToScreen()

        self.Shown += self.startProgress

    # Computation
    def startProgress(self, s, e):
        def update():
            
            for i in range(self.prog.Maximum):
                def step():
                    
                    # Insert Loop Start using int i
                    try:
                        if parameterToChange == "Model Group 'Name'":
                            areaNo = "%s" % (areas[i].GetParameters("Room-Apartment Number")[0].AsString())
                            for j in range(len(casework)):
                                caseNo = "%s" %(casework[j].GetParameters("Room-Apartment Number")[0].AsString())
                                if caseNo == areaNo:
                                    try:
                                        eleID = casework[j].GroupId
                                        ele = doc.GetElement(eleID)
                                        areas[i].GetParameters("Item No")[0].Set("%s" % (ele.Name))
                                        break
                                    except:
                                        continue

                        elif parameterToChange == "Area 'Name'":
                            areas[i].GetParameters("Item No")[0].Set(areas[i].GetParameters("Name")[0].AsString())
                    except:
                        None
                    # Loop End

                    self.prog.Value = i + 1
                    self.prog.Refresh()
                    percent = (int)(((float)(self.prog.Value) / (float)(self.prog.Maximum)) * 100)
                    self.prog.CreateGraphics().DrawString(("%s" % percent + "%"), Font("Arial", 8.25, FontStyle.Regular), Brushes.Black, PointF(190,3))

                self.Invoke(CallTarget0(step))

            if self.prog.Value == self.prog.Maximum:
                self.Close()
                    
        t = Thread(ThreadStart(update))
        t.Start()


# Transaction
with DB.Transaction(doc, 'Transfering Info to Area') as trans:
    trans.Start()

    # Start Progress Bar
    Application.EnableVisualStyles()
    f = ProgressBarDialog(len(areas), 'Transfering Info to Area')
    f.ShowDialog()

    trans.Commit()

# Data Collection

script.get_results().Calculations = len(areas)
script.get_results().Minutes = 3
