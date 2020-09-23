VERSION 5.00
Begin VB.UserControl dkSubClass 
   BackColor       =   &H000000FF&
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2145
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   2145
End
Attribute VB_Name = "dkSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Type RECT
'   Left As Long
'   Top As Long
'   Right As Long
'   Bottom As Long
'End Type
'
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Const SWP_FRAMECHANGED = &H20
'Private Const SWP_NOACTIVATE = &H10
'Private Const SWP_ALL = SWP_FRAMECHANGED Or SWP_NOACTIVATE
'
'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Private Const SW_SHOWNOACTIVATE = 4
'
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'
'Private Const BM_SETIMAGE As Long = &HF7
'Private Const BS_BITMAP = &H80
'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
'Private Const LR_LOADFROMFILE = &H10
'Private Const IMAGE_BITMAP = 0

Private WithEvents SubClass As clsSubclass
Attribute SubClass.VB_VarHelpID = -1
Private Started As Boolean
'Private ParenthWnd  As Long
'Private hWndButton As Long

Event Message(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'Default Property Values:
Const m_def_hWnd = 0
'Property Variables:
Dim m_hWnd As Long

Private Sub SubClass_Message(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    RaiseEvent Message(ByVal uMsg, ByVal wParam, ByVal lParam)
End Sub

Private Sub UserControl_Initialize()
Dim hBitmap As Long
    If Started Then
        If UserControl.Ambient.UserMode Then
            If m_hWnd = 0 Then Exit Sub
            
            Set SubClass = New clsSubclass
            SubClass.hwnd = m_hWnd
            Module1.Hook SubClass
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    hwnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
End Sub

Private Sub UserControl_Resize()
    Width = 420
    Height = 420
End Sub

Private Sub UserControl_Terminate()
    If Started Then
        Module1.UnHook SubClass
        Set SubClass = Nothing
    End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_hWnd = m_def_hWnd
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = m_hWnd
End Property

Public Property Let hwnd(ByVal New_hWnd As Long)
    m_hWnd = New_hWnd
    Started = m_hWnd <> 0
    UserControl_Initialize
    PropertyChanged "hWnd"
End Property

