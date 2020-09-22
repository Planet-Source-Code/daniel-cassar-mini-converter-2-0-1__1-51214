VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniConverter Options"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "Always On Top?"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkOnTop_Click()
    Dim lR As Long
    If chkOnTop.Value = 1 Then
        lR = SetTopMostWindow(frmMain.hwnd, True)
        gbOnTop = True
    Else
        lR = SetTopMostWindow(frmMain.hwnd, False)
        gbOnTop = False
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If gbOnTop = True Then chkOnTop.Value = 1
End Sub
