#pylint: disable=E0401,C0103,maybe-no-member

# Import pyRevit modules
from pyrevit import script
from pyrevit.framework import List
from pyrevit.framework import clr
from pyrevit import api
from pyrevit import revit, DB, HOST_APP

from Autodesk.Revit import ApplicationServices

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


__context__ = 'zerodoc'
__title__ = 'Test'
__doc__ = '.'
__author__ = 'Sebastian Teo'
__helpurl__ = 'mailto:steo@hayball.com.au'

filestring = "%s" % "C:/Users/steo/Documents/2357_Kaufland HQ_facade_carpark_A18_CENTRAL_steoYUFNP.rvt"

doc = ApplicationServices.Application().OpenDocumentFile(DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(filestring), DB.OpenOptions())

for a in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToElements():
    print("%s" % a)