#pylint: disable=E0401,C0103,maybe-no-member

import sys
import json
import ctypes
import os
import platform
import datetime

from pyrevit import HOST_APP, framework
from pyrevit import DB
from pyrevit import forms
from pyrevit import script
from pyrevit.loader import sessioninfo

def get_free_space_mb(dirname):
    """Return folder/drive free space (in gb)."""
    free_bytes = ctypes.c_ulonglong(0)
    ctypes.windll.kernel32.GetDiskFreeSpaceExW(ctypes.c_wchar_p(dirname), None, None, ctypes.pointer(free_bytes))
    return free_bytes.value / 1024 / 1024 / 1024

space = get_free_space_mb('C:')
msg = "Disk Space Warning\n\nYou have %sgb of Free Space on your drive, which is lower than the 20gb required for Revit to work properly.\n\nIf you continue without cleaning or increasing you Disk Space, you risk corrupting Central Revit file.\n\nPlease contact IT or the BIM Team if you need help." % (space)

if space < 20:
    forms.alert(msg, title='Hayball')

uuid = sessioninfo.get_session_uuid()
revitVersion = sessioninfo.get_runtime_info()[2]

fileName = "U:\\Revit Project Logs\\Log_%s_%s_%s.json" % (HOST_APP.username, revitVersion, uuid)

f = open(fileName, "w+")
f.close()

jsonArray = []
docopen = {}
docsave = {}
docsync = {}

# define event handler
def docopen_eventhandler(sender, args):

    path_str = args.PathName.replace("_%s" % HOST_APP.username, "")
    index = path_str.rfind("\\") + 1
    path_str = path_str[index:]
    dotindex = path_str.rfind(".") + 1
    filetype = path_str[dotindex:]

    docopen["File Path"] = "%s" % (path_str)
    docopen["File Type"] = "%s" % (filetype)
    docopen["Action"] = "Open Document"
    docopen["Date"] = datetime.datetime.now().strftime("%Y-%m-%d")
    docopen["Start Time"] = datetime.datetime.now().strftime("%H:%M:%S")
    docopen["User"] = "%s" % (HOST_APP.username)
    docopen["Location"] = args.PathName

def docopened_eventhandler(sender, args):
    
    docopen["Complete Time"] = datetime.datetime.now().strftime("%H:%M:%S")
    
    jsonArray.append(docopen)

    with open(fileName, "w") as f:
        json.dump(jsonArray, f)

def docsave_eventhandler(sender, args):
    
    path_str = args.Document.PathName.replace("_%s" % HOST_APP.username, "")
    index = path_str.rfind("\\") + 1
    path_str = path_str[index:]
    dotindex = path_str.rfind(".") + 1
    filetype = path_str[dotindex:]
    
    docsave["File Path"] = "%s" % (path_str)
    docsave["File Type"] = "%s" % (filetype)
    docsave["Action"] = "Save Document"
    docsave["Date"] = datetime.datetime.now().strftime("%Y-%m-%d")
    docsave["Start Time"] = datetime.datetime.now().strftime("%H:%M:%S")
    docsave["User"] = "%s" % (HOST_APP.username)
    docsave["Location"] = args.Document.PathName

def docsaved_eventhandler(sender, args):
    
    docsave["Complete Time"] = datetime.datetime.now().strftime("%H:%M:%S")
    
    jsonArray.append(docsave)

    with open(fileName, "w") as f:
        json.dump(jsonArray, f)

def docsync_eventhandler(sender, args):
    
    path_str = args.Document.PathName.replace("_%s" % HOST_APP.username, "")
    index = path_str.rfind("\\") + 1
    path_str = path_str[index:]
    dotindex = path_str.rfind(".") + 1
    filetype = path_str[dotindex:]
    
    docsync["File Path"] = "%s" % (path_str)
    docsync["File Type"] = "%s" % (filetype)
    docsync["Action"] = "Sync Document"
    docsync["Date"] = datetime.datetime.now().strftime("%Y-%m-%d")
    docsync["Start Time"] = datetime.datetime.now().strftime("%H:%M:%S")
    docsync["User"] = "%s" % (HOST_APP.username)
    docsync["Location"] = args.Document.PathName

def docsynced_eventhandler(sender, args):
    
    docsync["Complete Time"] = datetime.datetime.now().strftime("%H:%M:%S")
    
    jsonArray.append(docsync)

    with open(fileName, "w") as f:
        json.dump(jsonArray, f)

HOST_APP.app.DocumentOpening += \
    framework.EventHandler[DB.Events.DocumentOpeningEventArgs](
        docopen_eventhandler
        )

HOST_APP.app.DocumentOpened += \
    framework.EventHandler[DB.Events.DocumentOpenedEventArgs](
        docopened_eventhandler
        )

HOST_APP.app.DocumentSaving += \
    framework.EventHandler[DB.Events.DocumentSavingEventArgs](
        docsave_eventhandler
        )

HOST_APP.app.DocumentSaved += \
    framework.EventHandler[DB.Events.DocumentSavedEventArgs](
        docsaved_eventhandler
        )

HOST_APP.app.DocumentSynchronizingWithCentral += \
    framework.EventHandler[DB.Events.DocumentSynchronizingWithCentralEventArgs](
        docsync_eventhandler
        )

HOST_APP.app.DocumentSynchronizedWithCentral += \
    framework.EventHandler[DB.Events.DocumentSynchronizedWithCentralEventArgs](
        docsynced_eventhandler
        )