VERSION 5.00
Object = "*\AdkSubClassProj.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin dkSubClassProj.dkSubClass dkSubClass1 
      Left            =   1320
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dkSubClass1_Message(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Debug.Print "Got message " & uMsg & " with " & wParam & ", " & lParam
End Sub

Private Sub Form_Load()
    dkSubClass1.hWnd = Me.hWnd
End Sub
