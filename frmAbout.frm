VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Daniel's Mini Converter"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   180
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   10
      Top             =   90
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   45
      TabIndex        =   1
      Top             =   645
      Width           =   4125
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   330
         Left            =   3240
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "2001 - 2003 - Daniel J. Cassar"
         Height          =   285
         Left            =   945
         TabIndex        =   9
         Top             =   945
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright:"
         Height          =   285
         Left            =   15
         TabIndex        =   8
         Top             =   945
         Width           =   870
      End
      Begin VB.Label Label7 
         Caption         =   "Daniel J. Cassar"
         Height          =   285
         Left            =   945
         TabIndex        =   7
         Top             =   690
         Width           =   2130
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Author:"
         Height          =   285
         Left            =   45
         TabIndex        =   6
         Top             =   690
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "Windows 9x or greater"
         Height          =   285
         Left            =   945
         TabIndex        =   5
         Top             =   435
         Width           =   2025
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Platform:"
         Height          =   285
         Left            =   30
         TabIndex        =   4
         Top             =   435
         Width           =   855
      End
      Begin VB.Label lblVersion 
         Height          =   285
         Left            =   945
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Version:"
         Height          =   285
         Left            =   210
         TabIndex        =   2
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MiniConverter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Close Window
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub
