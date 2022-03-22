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
__title__ = 'Number Carparks\nwith Paths'
__doc__ = "Assigns Carpark numbers automatically using the 'Label Maker' line style as the path of order. Deletes 'Label Maker' lines at completion."
__author__ = 'Sebastian Teo'
__helpurl__ = 'mailto:steo@hayball.com.au'

# Get Current Revit Document
doc = revit.doc
uiapp = HOST_APP.uiapp

# Inputs
counter = 1
prefix = ""
suffix = ""

# Form Class
class IForm(Form):
    
    tx1 = TextBox(Text='1')
    tx2 = TextBox(Text='')
    tx3 = TextBox(Text='')
    
    def __init__(self):

        # Logo
        logo = Panel()
        self.Controls.Add(logo)
        img = Image.FromFile("%s" % os.path.dirname(os.path.realpath(__file__)) + "\\Hayball_Logo.png")
        logo.BackgroundImage = img
        logo.Location = Point(8, 220)
        logo.Size = Size(200,50)
        
        self.Size = Size(350, 320)
        self.Text = "Select Options"

        head = Label(Text = 'Options for automatically numbering carpark bays using the "Label Maker" Line style. Lines will be deleted after completion')
        self.Controls.Add(head)
        head.Size = Size(200,50)
        head.Location = Point(10,8)

        lb = Label(Text = 'Number to start count from:')
        self.Controls.Add(lb)
        lb.Size = Size(200,18)
        lb.Location = Point(10,65)
        
        lb1 = Label(Text = 'Prefix (Leave Blank for None):')
        self.Controls.Add(lb1)
        lb1.Size = Size(200,18)
        lb1.Location = Point(10,110)
        
        lb2 = Label(Text = 'Suffix (Leave Blank for None):')
        self.Controls.Add(lb2)
        lb2.Size = Size(200,18)
        lb2.Location = Point(10,155)
        
        self.tx1.Parent = self
        self.tx1.Size = Size(200,25)
        self.tx1.Location = Point(10, 85)
        
        self.tx2.Parent = self
        self.tx2.Size = Size(200,25)
        self.tx2.Location = Point(10, 130)

        self.tx3.Parent = self
        self.tx3.Size = Size(200,25)
        self.tx3.Location = Point(10, 175)

        self.ControlBox = True

        done = Button(Text='Done')
        self.Controls.Add(done)
        done.Location = Point(235, 15)
        done.Click += self.doneClick
        
        self.CenterToScreen()
            
    def doneClick(self, sender, event):
        
        global gcounter
        global gprefix
        global gsuffix

        gprefix = self.tx2.Text
        gsuffix = self.tx3.Text

        try:
            gcounter = int(self.tx1.Text)
        except:
            print("Inputs must be Whole Numbers")
            self.Close()
        
        self.Close()

IForm().ShowDialog()

if 'gprefix' in globals():
    prefix = gprefix
else:
    sys.exit()

if 'gsuffix' in globals():
    suffix = gsuffix
else:
    sys.exit()

try:
    counter = gcounter
except:
    sys.exit()

labelPath = []

collector = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(DB.ElementClassFilter(DB.CurveElement))
for i in collector:
    if i.LineStyle.Name == 'Label Maker':
        labelPath.append(i.GeometryCurve)

# Polyline
tol = (1.0/12.0)/16.0

ptsList = []
ptsListClean = []
temp = labelPath[:]

ptsList.append(temp[0].GetEndPoint(0))
buf = temp[0].GetEndPoint(1)
del temp[0]

for a in range(len(labelPath) - 1):
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

ptsList.append(buf)

for c in range(len(ptsList)):
    if c != (len(ptsList)-1):
        if ptsList[c].DistanceTo(ptsList[c+1]) > doc.Application.ShortCurveTolerance:
            ptsListClean.append(ptsList[c])
    else:
        if ptsList[c].DistanceTo(ptsList[0]) > doc.Application.ShortCurveTolerance:
            ptsListClean.append(ptsList[c])

for p in range(len(ptsListClean)):
    ptsListClean[p] = DB.XYZ(ptsListClean[p].X, ptsListClean[p].Y, 0)

poly = DB.PolyLine.Create(ptsListClean)
length = 0
for j in labelPath:
    length += j.ApproximateLength

spacing = DB.UnitUtils.ConvertToInternalUnits(500, DB.DisplayUnitType.DUT_MILLIMETERS)
spacingNo = int(length // spacing) + 1
boundedSpacing = 1.0 / spacingNo

ptList = []

for i in range(spacingNo):
    para = i * boundedSpacing
    pt = poly.Evaluate(para)
    ptList.append(pt)

parking = DB.FilteredElementCollector(doc,doc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_Parking).WhereElementIsNotElementType().ToElements()
parkingBounds = []

for el in parking:
    bounds = []
    BB = el.get_BoundingBox(doc.ActiveView)
    bounds.append(DB.XYZ(BB.Min.X, BB.Min.Y, 0))
    bounds.append(DB.XYZ(BB.Max.X, BB.Max.Y, 0))
    parkingBounds.append(bounds)

orderedElements = []

for xy in ptList:
    for alpha in range(len(parking)):
        if parking[alpha] not in orderedElements:
            if (xy.X > parkingBounds[alpha][0].X and xy.X < parkingBounds[alpha][1].X):
                if (xy.Y > parkingBounds[alpha][0].Y and xy.Y < parkingBounds[alpha][1].Y):
                    orderedElements.append(parking[alpha])
                    break

# Check
user = HOST_APP.username
elementsToCheck = orderedElements

for element in elementsToCheck:
    bufstr = clr.Reference[str]()
    DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id, bufstr)
    _bufstr = '%s' % bufstr
    if _bufstr != "''":
        if _bufstr != "'%s'" % user:
            print("%s owns an element that will be modified. Please request release of element before trying again" % _bufstr)
            sys.exit()

for element in collector:
    if element.LineStyle.Name == 'Label Maker':
        bufstr = clr.Reference[str]()
        DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id, bufstr)
        _bufstr = '%s' % bufstr
        if _bufstr != "''":
            if _bufstr != "'%s'" % user:
                print("%s owns an element that will be modified. Please request release of element before trying again" % _bufstr)
                sys.exit()

#Transaction

with DB.Transaction(doc, "Numbering Carpark") as trans:
    
    trans.Start()

    for thing in orderedElements:
        if counter < 10:
            thing.GetParameters("Mark")[0].Set(prefix + "0" + "%s" % counter + suffix)
            counter += 1
        else:
            thing.GetParameters("Mark")[0].Set(prefix + "%s" % counter + suffix)
            counter += 1

    for i in collector:
        if i.LineStyle.Name == 'Label Maker':
            doc.Delete(i.Id)

    trans.Commit()

# Data Collection

script.get_results().Calculations = len(orderedElements)
script.get_results().Minutes = 3
