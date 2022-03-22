#pylint: disable=E0401,C0103,maybe-no-member

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit.framework import List
from pyrevit.framework import clr
from pyrevit import revit, DB, HOST_APP, UI

import ctypes
from ctypes import wintypes

# Import Winform Modules
clr.AddReference("System.Windows.Forms")
clr.AddReference("system.Drawing")
clr.AddReference('IronPython')
clr.AddReference("Microsoft.Office.Interop.Outlook")
from System.Windows.Forms import Panel, GroupBox, RadioButton, ProgressBar, Application, Form, CheckBox, CheckedListBox, ListBox, Button, Label, TextBox
from System.Drawing import PointF, Size, Point, Font, FontStyle, Brushes, Image
from System.Threading import ThreadStart, Thread
from IronPython.Compiler import CallTarget0
from System.Runtime.InteropServices import Marshal
from Microsoft.Office.Interop.Outlook import OlWindowState

# Import System modules
import System
import sys
import os
from System.Collections.Generic import *

__title__ = 'Email for Help\nor Suggestions'
__author__ = 'Sebastian Teo'

try:
    mail= Marshal.GetActiveObject("Outlook.Application").CreateItem(0)
    
except:
    forms.alert("Please Open Outlook for this button to work!", title="Email")
    sys.exit()

inspector = mail.GetInspector
outlook = Marshal.GetActiveObject("Outlook.Application")
outlook.ActiveExplorer().WindowState = OlWindowState.olMinimized
outlook.ActiveExplorer().WindowState = OlWindowState.olMaximized
mail.Recipients.Add("steo@hayball.com.au")
mail.Recipients.Add("erajapackiyam@hayball.com.au")
mail.Subject = "Revit - Hayball Plugin"
mail.Body = "Hi Beautiful people! Here is my query below: \n\n"
mail.Display(False)
inspector.Activate()

wintypes.ULONG_PTR = wintypes.WPARAM
user32 = ctypes.WinDLL('user32', use_last_error=True)

INPUT_MOUSE    = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARE = 2

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002
KEYEVENTF_UNICODE     = 0x0004
KEYEVENTF_SCANCODE    = 0x0008

MAPVK_VK_TO_VSC = 0

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx",          wintypes.LONG),
                ("dy",          wintypes.LONG),
                ("mouseData",   wintypes.DWORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk",         wintypes.WORD),
                ("wScan",       wintypes.WORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

    def __init__(self, *args, **kwds):
        super(KEYBDINPUT, self).__init__(*args, **kwds)
        # some programs use the scan code even if KEYEVENTF_SCANCODE
        # isn't set in dwFflags, so attempt to map the correct code.
        if not self.dwFlags & KEYEVENTF_UNICODE:
            self.wScan = user32.MapVirtualKeyExW(self.wVk,
                                                 MAPVK_VK_TO_VSC, 0)

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg",    wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = (("ki", KEYBDINPUT),
                    ("mi", MOUSEINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("_input",)
    _fields_ = (("type",   wintypes.DWORD),
                ("_input", _INPUT))

LPINPUT = ctypes.POINTER(INPUT)

def _check_count(result, func, args):
    if result == 0:
        raise ctypes.WinError(ctypes.get_last_error())
    return args

user32.SendInput.errcheck = _check_count
user32.SendInput.argtypes = (wintypes.UINT, # nInputs
                             LPINPUT,       # pInputs
                             ctypes.c_int)  # cbSize

def PressKey(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

PressKey(0x28)

sys.exit()
