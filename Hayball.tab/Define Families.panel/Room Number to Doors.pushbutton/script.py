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
__title__ = 'Room Info To\nDoor'
__doc__ = 'Moves the Room-Apartment number from the area into all Door families which are contained within that room.'
__author__ = 'Sebastian Teo'
__helpurl__ = 'mailto:steo@hayball.com.au'

# Get Current Revit Document
doc = revit.doc

# Inputs
l =[]
elementsToCheck = []

for i in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements():
    if ("%s" % (i.GetType())) == "Autodesk.Revit.DB.FamilyInstance":
        elementsToCheck.append(i)

for para in elementsToCheck[0].Parameters:
    if para.UserModifiable:
        if 'String' in para.StorageType.ToString():
            l.append("%s" % para.Definition.Name)

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
        lab = Label(Text = 'Select Parameter to transfer to')
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
    if returnList != "None":
        parameterToChange = returnList
    else:
        sys.exit()
except:
    sys.exit()

#areasName = []
#areasGeo = []
#
#rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
#
#for room in rooms:
#    if room.GetParameters("Room-Apartment Number")[0].AsString() != None:
#        for solid in room.ClosedShell:
#            areasGeo.append(solid)
#            areasName.append(room.GetParameters("Room-Apartment Number")[0].AsString())
            
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
                        room = elementsToCheck[i].ToRoom[phase]
                        if room != None:
                            name = room.GetParameters("Room-Apartment Number")[0].AsString()
                            if name != None:
                                elementsToCheck[i].GetParameters(parameterToChange)[0].Set(name)
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


# Check
user = HOST_APP.username

for element in elementsToCheck:
    bufstr = clr.Reference[str]()
    DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id, bufstr)
    _bufstr = '%s' % bufstr
    if _bufstr != "''":
        if _bufstr != "'%s'" % user:
            print("%s owns an element that will be modified. Please request release of element before trying again" % _bufstr)
            sys.exit()

phase = list(doc.Phases)[-1]

#Transaction
with DB.Transaction(doc, 'Room to Door') as trans:
    trans.Start()
    
    # Start Progress Bar
    Application.EnableVisualStyles()
    f = ProgressBarDialog(len(elementsToCheck), 'Transfering Parameters to Door')
    f.ShowDialog()
        
    trans.Commit()

# Data Collection

script.get_results().Calculations = len(elementsToCheck)
script.get_results().Minutes = 3

# Audit