#pylint: disable=E0401,C0103,maybe-no-member

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
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

__title__ = 'Who did that??'
__author__ = 'Ehsan Iran-Nejad\n'\
             'Frederic Beaupere\n'\
             'Sebastian Teo'


def who_reloaded_keynotes():
    location = revit.doc.PathName
    if revit.doc.IsWorkshared:
        try:
            modelPath = \
                DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(location)
            transData = DB.TransmissionData.ReadTransmissionData(modelPath)
            externalReferences = transData.GetAllExternalFileReferenceIds()
            for refId in externalReferences:
                extRef = transData.GetLastSavedReferenceData(refId)
                path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(
                    extRef.GetPath()
                    )

                if extRef.ExternalFileReferenceType == \
                        DB.ExternalFileReferenceType.KeynoteTable \
                        and '' != path:
                    ktable = revit.doc.GetElement(extRef.GetReferencingId())
                    editedByParam = \
                        ktable.Parameter[DB.BuiltInParameter.EDITED_BY]
                    if editedByParam and editedByParam.AsString() != '':
                        forms.alert('Keynote table has been reloaded by:\n'
                                    '{0}\nTable Id is: {1}'
                                    .format(editedByParam.AsString(),
                                            ktable.Id))
                    else:
                        forms.alert('No one own the keynote table. '
                                    'You can make changes and reload.\n'
                                    'Table Id is: {0}'.format(ktable.Id))
        except Exception as e:
            forms.alert('Model is not saved yet. '
                        'Can not aquire keynote file location.')
    else:
        forms.alert('Model is not workshared.')


def who_created_selection():
    selection = revit.get_selection()
    if revit.doc.IsWorkshared:
        if selection and len(selection) == 1:
            eh = revit.query.get_history(selection.first)

            forms.alert('Creator: {0}\n'
                        'Current Owner: {1}\n'
                        'Last Changed By: {2}'.format(eh.creator,
                                                      eh.owner,
                                                      eh.last_changed_by))
        else:
            forms.alert('Exactly one element must be selected.')
    else:
        forms.alert('Model is not workshared.')


def who_created_activeview():
    active_view = revit.activeview
    view_id = active_view.Id.ToString()
    view_name = active_view.Name
    view_creator = \
        DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc,
                                                      active_view.Id).Creator

    forms.alert('{}\n{}\n{}'
                .format("Creator of the current view: " + view_name,
                        "with the id: " + view_id,
                        "is: " + view_creator))


options = {'Who Created Active View?': who_created_activeview,
           'Who Created Selected Element?': who_created_selection,
           'Who Reloaded Keynotes Last?': who_reloaded_keynotes}

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
        lab = Label(Text = 'Who...')
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
IForm(options.keys()).ShowDialog()

try:
    if returnList != "None":
        selected_option = returnList
    else:
        sys.exit()
except:
    sys.exit()

if selected_option:
    options[selected_option]()
