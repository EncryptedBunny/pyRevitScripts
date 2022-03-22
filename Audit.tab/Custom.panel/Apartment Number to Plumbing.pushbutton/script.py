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
__title__ = 'Area Info To\nPlumbing'
__doc__ = 'Moves the Room-Apartment number from the area into all Plumbing Fixtures families which are contained within that area.'
__author__ = 'Sebastian Teo'
__helpurl__ = 'mailto:steo@hayball.com.au'

# Get Current Revit Document
doc = revit.doc

# Inputs
inputList = ['Room-Apartment Number', 'Item No']

# Radio Button Class
class IForm(Form):
    
    gb = GroupBox()
    selected = RadioButton()
    tx2 = TextBox(Text='1500')
    tx1 = TextBox(Text='0')
    
    def __init__(self, dataEnteringNode):

        # Logo
        logo = Panel()
        self.Controls.Add(logo)
        img = Image.FromFile("%s" % os.path.dirname(os.path.realpath(__file__)) + "\\Hayball_Logo.png")
        logo.BackgroundImage = img
        logo.Location = Point(10, 220)
        logo.Size = Size(200,50)
        
        self.Size = Size(350, 320)
        self.Text = "Select Options"
        self.gb.Parent = self
        self.gb.Size = Size(200,65)
        self.gb.Location = Point(10,35)

        lb = Label(Text = 'Select Parameter to Assign to Plumbing Fixtures')
        self.Controls.Add(lb)
        lb.Size = Size(200,30)
        lb.Location = Point(10,7)
        
        lb1 = Label(Text = 'Offset off Area Level (mm)')
        self.Controls.Add(lb1)
        lb1.Size = Size(200,18)
        lb1.Location = Point(10,110)
        
        lb2 = Label(Text = 'Search Height (mm)')
        self.Controls.Add(lb2)
        lb2.Size = Size(200,18)
        lb2.Location = Point(10,155)
        
        self.tx1.Parent = self
        self.tx1.Size = Size(200,25)
        self.tx1.Location = Point(10, 130)
        
        self.tx2.Parent = self
        self.tx2.Size = Size(200,25)
        self.tx2.Location = Point(10, 175)

        self.ControlBox = True

        rb1 = RadioButton(Text='%s' % dataEnteringNode[0])
        self.gb.Controls.Add(rb1)
        rb1.Location = Point(10,15)
        rb1.Size = Size(155,17)
        rb1.Checked = True
        self.selected = rb1
        rb1.CheckedChanged += self.radioClick

        rb2 = RadioButton(Text='%s' % dataEnteringNode[1])
        self.gb.Controls.Add(rb2)
        rb2.Location = Point(10,35)
        rb2.Size = Size(155,17)
        rb2.CheckedChanged += self.radioClick
     
        done = Button(Text='Done')
        self.Controls.Add(done)
        done.Location = Point(235, 15)
        done.Click += self.doneClick
        
        self.CenterToScreen()

    def radioClick(self, sender, event):
        rb = sender
        if rb.Checked:
            self.selected = rb
            
    def doneClick(self, sender, event):
        
        global returnList
        global offsetNo
        global extrusionNo
        returnList = "%s" % self.selected.Text
        try:
            offsetNo = int(self.tx1.Text)
            extrusionNo = int(self.tx2.Text)
        except:
            print("Inputs must be Whole Numbers")
            self.Close()
        
        self.Close()

IForm(inputList).ShowDialog()

if 'offsetNo' in globals():
    offset = offsetNo
else:
    sys.exit()

if 'extrusionNo' in globals():
    extrusionDis = extrusionNo
else:
    sys.exit()

try:
    parameterName = returnList
except:
    sys.exit()

# Define Area names to collect
filterList = ["1B", "2B", "3B", "TH", "STUDIO", "TWIN", "CLUSTER"]

areas = []
areasName = []
areasLns = []
areasGeo = []

# Background Collection
for a in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToElements():
    for filter in filterList:
        if a.GetParameters("Area")[0].AsDouble() != 0:
            apt = a.GetParameters(parameterName)[0].AsString()
            if apt != "":
                buf = []
                for i in a.GetBoundarySegments(DB.SpatialElementBoundaryOptions())[0]:
                    buf.append(i.GetCurve())
                areas.append(buf)
                areasName.append(apt)

tol = (1.0/12.0)/16.0

for i in range(len(areas)):
    
    ptsList = []
    ptsListClean = []
    temp = areas[i][:]
    crvLoop = []
    
    ptsList.append(temp[0].GetEndPoint(0))
    buf = temp[0].GetEndPoint(1)
    del temp[0]
    
    for a in range(len(areas[i]) - 1):
        for cv in temp:
            
            if cv.GetEndPoint(0).DistanceTo(buf) < tol:
                ptsList.append(cv.GetEndPoint(0))
                buf = cv.GetEndPoint(1)
                del temp[temp.index(cv)]
        
            elif cv.GetEndPoint(1).DistanceTo(buf) < tol:
                ptsList.append(cv.GetEndPoint(1))
                buf = cv.GetEndPoint(0)
                del temp[temp.index(cv)]
                
            else:
                dis = []
                for closest in temp:
                    dis.append(closest.GetEndPoint(0).DistanceTo(buf))
                    dis.append(closest.GetEndPoint(1).DistanceTo(buf))
                
                bufline = temp[(dis.index(min(dis)) // 2)]
                closept = dis.index(min(dis)) % 2
                
                ptsList.append(bufline.GetEndPoint(closept))
                buf = bufline.GetEndPoint(((closept + 1) % 2))
                del temp[temp.index(bufline)]
    
    for c in range(len(ptsList)):
        if c != (len(ptsList)-1):
            if ptsList[c].DistanceTo(ptsList[c+1]) > doc.Application.ShortCurveTolerance:
                ptsListClean.append(ptsList[c])
        else:
            if ptsList[c].DistanceTo(ptsList[0]) > doc.Application.ShortCurveTolerance:
                ptsListClean.append(ptsList[c])
    
    zmove = (ptsListClean[0].Z + DB.UnitUtils.ConvertToInternalUnits(offset, DB.DisplayUnitType.DUT_MILLIMETERS))
    for p in range(len(ptsListClean)):
        ptsListClean[p] = DB.XYZ(ptsListClean[p].X, ptsListClean[p].Y, zmove)
    
    for k in range(len(ptsListClean)):
        if k != (len(ptsListClean)-1):
            ln = DB.Line.CreateBound(ptsListClean[k], ptsListClean[k+1])
        else:
            ln = DB.Line.CreateBound(ptsListClean[k], ptsListClean[0])
        crvLoop.append(ln)
    
    crvLoopFam = DB.CurveLoop.Create(crvLoop)
    areasLns.append(crvLoopFam)

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
                        collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType()
                        plumbing = (collector.WherePasses(DB.ElementIntersectsSolidFilter(areasGeo[i])).ToElements())
                        for item in plumbing:
                            item.GetParameters(parameterName)[0].Set(areasName[validIndex[i]])
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

validIndex = []
invalidRooms = []

for CL in areasLns:
    try:
        l = []
        l.Add(CL)
        if CL.HasPlane():
            geo = DB.GeometryCreationUtilities.CreateExtrusionGeometry(l, DB.XYZ(0,0,1), DB.UnitUtils.ConvertToInternalUnits(extrusionDis, DB.DisplayUnitType.DUT_MILLIMETERS))
            areasGeo.append(geo)
            validIndex.append(areasLns.index(CL))
        else:
            invalidRooms.append(areasName[areasLns.index(CL)])
    except:
        invalidRooms.append(areasName[areasLns.index(CL)])

# Check
user = HOST_APP.username
elementsToCheck = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements()

for element in elementsToCheck:
    bufstr = clr.Reference[str]()
    DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id, bufstr)
    _bufstr = '%s' % bufstr
    if _bufstr != "''":
        if _bufstr != "'%s'" % user:
            print("%s owns an element that will be modified. Please request release of element before trying again" % _bufstr)
            sys.exit()

#Transaction
with DB.Transaction(doc, 'Area to Plumbing Fixtures') as trans:
    trans.Start()
    
    # Start Progress Bar
    Application.EnableVisualStyles()
    f = ProgressBarDialog(len(areasGeo), 'Transfering Parameters to Plumbing Fixtures')
    f.ShowDialog()
        
    trans.Commit()

# Data Collection

script.get_results().Calculations = len(areasGeo)
script.get_results().Minutes = 3

# Audit
invalidRooms.sort()