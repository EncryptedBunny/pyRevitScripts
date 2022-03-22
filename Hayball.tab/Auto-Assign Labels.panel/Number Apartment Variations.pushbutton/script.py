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
from System.Windows.Forms import GroupBox, RadioButton, ProgressBar, Application, Form, CheckBox, CheckedListBox, ListBox, Button, Label, TextBox
from System.Drawing import PointF, Size, Point, Font, FontStyle, Brushes
from System.Threading import ThreadStart, Thread
from IronPython.Compiler import CallTarget0

# Import System modules
import System
import sys
from System.Collections.Generic import *

# Set component Metadata
__title__ = 'Apartment\nVariations'
__doc__ = 'Automatically Increments Apartment type Variations eg. "1B1B" is prefixed with "_1A", "_1B", etc according to variations'
__author__ = 'Sebastian Teo'
__helpurl__ = 'mailto:steo@hayball.com.au'

# Get Current Revit Document
doc = revit.doc

# Functions
def increment_char(c):
    """
    Increment an uppercase character, returning 'A' if 'Z' is given
    """
    return chr(ord(c) + 1) if c != 'Z' else 'A'
    
def increment_str(s):
    lpart = s.rstrip('Z')
    if not lpart:  # s contains only 'Z'
        new_s = 'A' * (len(s) + 1)
    else:
        num_replacements = len(s) - len(lpart)
        new_s = lpart[:-1] + increment_char(lpart[-1])
        new_s += 'A' * num_replacements
    return new_s


# Code
uniqueNamesExcluded = []
uniqueNames = []
names = []

groups = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_IOSModelGroups).WhereElementIsNotElementType().ToElements()

for search in groups:
    names.append(search.Name)
    if search.Name not in uniqueNamesExcluded:
        if "(members excluded)" not in search.Name:
            uniqueNamesExcluded.append(search.Name)
            
terms = ['1B1BS','2B1BS','2B2BS','3B2BS','3B3BS','1B1B_TH','2B1B_TH','2B2B_TH','3B2B_TH','3B3B_TH','1B1B','2B1B','2B2B','3B2B','3B3B','STUDIO','TWIN','CLUSTERS']

counters = ['A'] * len(terms)

uniqueNamesExcluded.sort()

# Check
user = HOST_APP.username
elementsToCheck = groups

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
                    
                    a = uniqueNamesExcluded[i]
                    index = names.index(a)
                    str = a

                    for j in range(len(counters)):

                        if terms[j][0] == 'S':
                            prefix = 'STUDIO'

                        elif terms[j][0] == 'T':
                            prefix = 'TWIN'

                        elif terms[j][0] == 'C':
                            prefix = 'CLUSTERS'

                        else:
                            prefix = terms[j][0]
                        
                        if terms[j] in str:
                            newstr = str.split(terms[j])[0] + terms[j] + '_' + prefix + counters[j]
                            counters[j] = increment_str(counters[j])
                            groups[index].GroupType.Name = newstr
                            break

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
with DB.Transaction(doc, 'Increment') as trans:
    trans.Start()

    # Start Progress Bar
    Application.EnableVisualStyles()
    f = ProgressBarDialog(len(uniqueNamesExcluded), 'Assigning Apartment Variations')
    f.ShowDialog()

    trans.Commit()

for search in groups:
    if search.Name not in uniqueNames:
        uniqueNames.append(search.Name)

# Data Collection
script.get_results().Calculations = len(uniqueNamesExcluded)
script.get_results().Minutes = 3

# Audit
uniqueNames.sort()
