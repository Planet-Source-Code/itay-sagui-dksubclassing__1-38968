VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarhWnd As Long 'local copy
Private mvarButtonhWnd As Long 'local copy

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent Message[(arg1, arg2, ... , argn)]
Public Event Message(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Friend Sub fRaiseEvent(ByVal EventID As Long, ByVal wParam As Long, ByVal lParam As Long)
    RaiseEvent Message(EventID, wParam, lParam)
End Sub


Public Property Let hwnd(ByVal vData As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.hWnd = 5
    mvarhWnd = vData
End Property

Public Property Get hwnd() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.hWnd
    hwnd = mvarhWnd
End Property
