VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MiniConverter"
   ClientHeight    =   2175
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3060
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboFrom 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   855
      TabIndex        =   2
      Top             =   900
      Width           =   2100
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Left            =   855
      TabIndex        =   1
      Top             =   465
      Width           =   2100
   End
   Begin VB.ComboBox cboTo 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   855
      TabIndex        =   3
      Top             =   1335
      Width           =   2100
   End
   Begin VB.TextBox txtTo 
      BackColor       =   &H80000000&
      Height          =   300
      Left            =   855
      TabIndex        =   4
      Top             =   1770
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   -360
      TabIndex        =   0
      Top             =   -60
      Width           =   3765
      Begin VB.Label lblCategory 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   1320
         TabIndex        =   5
         Top             =   140
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   " Category:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   360
         TabIndex        =   6
         Top             =   135
         Width           =   1020
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      Height          =   255
      Left            =   75
      TabIndex        =   10
      Top             =   945
      Width           =   705
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Convert:"
      Height          =   285
      Left            =   75
      TabIndex        =   9
      Top             =   495
      Width           =   705
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Answer:"
      Height          =   225
      Left            =   75
      TabIndex        =   8
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   1380
      Width           =   705
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPrecisionT 
         Caption         =   "&Precision"
         Begin VB.Menu mnuPrecision 
            Caption         =   "0 Places"
            Index           =   0
         End
         Begin VB.Menu mnuPrecision 
            Caption         =   "1 Places"
            Index           =   1
         End
         Begin VB.Menu mnuPrecision 
            Caption         =   "2 Places"
            Index           =   2
         End
         Begin VB.Menu mnuPrecision 
            Caption         =   "3 Places"
            Index           =   3
         End
         Begin VB.Menu mnuPrecision 
            Caption         =   "6 Places"
            Index           =   6
         End
         Begin VB.Menu mnuPrecision 
            Caption         =   "9 Places"
            Checked         =   -1  'True
            Index           =   9
         End
      End
      Begin VB.Menu mnuFavourite 
         Caption         =   "&Favourite"
         Begin VB.Menu mnuFav 
            Caption         =   "Get Favourite 1"
            Index           =   1
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuFav 
            Caption         =   "Get Favourite 2"
            Index           =   2
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuFav 
            Caption         =   "Get Favourite 3"
            Index           =   3
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuspace4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSetFav 
            Caption         =   "Set Favourite 1"
            Index           =   1
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuSetFav 
            Caption         =   "Set Favourite 2"
            Index           =   2
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuSetFav 
            Caption         =   "Set Favourite 3"
            Index           =   3
            Shortcut        =   ^{F4}
         End
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuCategory 
      Caption         =   "&Category"
      Begin VB.Menu MnuEdit 
         Caption         =   "Edit..."
      End
      Begin VB.Menu mnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCategoryName 
         Caption         =   "CategoryName"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Mini Converter"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Categorys As clsCategorys
Dim Reg As CRegSettings

Private Sub Form_Load()
    gbPrecision = 9                     ' Set Display Precision to 9
    gbDataPath = App.Path & "\data\"    ' Set Category Datafiles Directory
    
    Set Categorys = New clsCategorys    'Load Categorys into Menu
    Categorys.loadCategoryData
    loadCategorys
    
    loadUnits cboFrom               ' Load Units into Combos
    loadUnits cboTo
End Sub

' Load Categorys into menu
Sub loadCategorys()
    Dim Category As clsCategory
    Dim i As Integer
    
    ' Clear the menu of categorys
    For i = 1 To (mnuCategoryName.Count - 1)
        Unload mnuCategoryName(i)
    Next i
    
    ' Fill Category Menu with categorys
    i = 0
    For Each Category In Categorys
        i = i + 1
        Load mnuCategoryName(i)
        mnuCategoryName(i).Caption = Category.strName
        mnuCategoryName(i).Visible = True
    Next Category

    ' If menu isn't blank goto first one
    If i > 0 Then
        lblCategory.Caption = mnuCategoryName(1).Caption
        mnuCategoryName(1).Checked = True
    End If

End Sub

' Load the category Units
Sub loadUnits(cbo As ComboBox)
    Dim Unit As clsUnit
    
    ' Clear the existing combobox and fill with category units
    cbo.Clear
    For Each Unit In Categorys(lblCategory.Caption).Units
        cbo.AddItem Unit.strName
    Next Unit
    
    ' Assign cboFrom and cboTo different Units
    If cbo.Name = "cboTo" And cbo.ListCount > 1 Then
        cbo.ListIndex = IIf(cbo.ListCount < 1, -1, 1)
    ElseIf cbo.ListCount > 0 Then
        cbo.ListIndex = IIf(cbo.ListCount < 0, -1, 0)
    End If
End Sub

'User chooses a differant Unit to convert from
Private Sub cboFrom_Click()
    CalculateAnswer
End Sub

'User chooses a differant Unit to convert to
Private Sub cboTo_Click()
    CalculateAnswer
End Sub

'Show About MiniConverter Dialog
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

' User clicks on a category
Private Sub mnuCategoryName_Click(Index As Integer)
    Dim i As Integer
    
    ' Uncheck other menu's and check the new selection
    For i = 1 To (mnuCategoryName.Count - 1)
        mnuCategoryName(i).Checked = False
    Next i
    mnuCategoryName(Index).Checked = True
    lblCategory.Caption = mnuCategoryName(Index).Caption
    
    cboTo.Text = "" ' To stop calculate answer from firing
    ' Fill combos with the categorys units and recalculate answer
    loadUnits cboFrom
    loadUnits cboTo
End Sub

' User wants to edit categories
Private Sub MnuEdit_Click()
    frmEdit.Show vbModal
    
    'Clear old values.
    Set Categorys = Nothing
    
    'Load new ones.
    Set Categorys = New clsCategorys
    Categorys.loadCategoryData
    
    loadCategorys
End Sub

' Quit Program
Private Sub mnuExit_Click()
    Unload Me
End Sub

' Get favourite category and units from the registry
Private Sub mnuFav_Click(Index As Integer)
    Dim strTemp As String
    Dim intfrom As String
    Dim i As Integer
        
    ' Define Registry Settings
    Set Reg = New CRegSettings
    Reg.Company = App.CompanyName
    Reg.AppName = APP_NAME
    
    ' Get favourite category and units from the registry
    strTemp = Reg.GetSetting("Favourite " & Index, "Category", "")
    
    ' Check if a cateogry was returned then check to see if it still exists in
    ' the data directory then select the category and select their units
    If Not strTemp = "" Then
        For i = 1 To (mnuCategoryName.Count - 1)
            If mnuCategoryName(i).Caption = strTemp Then
                mnuCategoryName_Click (i)
                cboFrom.ListIndex = Reg.GetSetting("Favourite " & Index, "From", "")
                cboTo.ListIndex = Reg.GetSetting("Favourite " & Index, "To", "")
            End If
        Next i
    End If
End Sub

' Set the precision
Private Sub mnuPrecision_Click(Index As Integer)
    gbPrecision = Index
    
    mnuPrecision(0).Checked = False
    mnuPrecision(1).Checked = False
    mnuPrecision(2).Checked = False
    mnuPrecision(3).Checked = False
    mnuPrecision(6).Checked = False
    mnuPrecision(9).Checked = False
    mnuPrecision(Index).Checked = True

    CalculateAnswer
End Sub

' Set favourite category and units from the registry
Private Sub mnuSetFav_Click(Index As Integer)
    ' Define Registry Settings
    Set Reg = New CRegSettings
    Reg.Company = App.CompanyName
    Reg.AppName = APP_NAME
    
    ' Store current selection
    Reg.SaveSetting "Favourite " & Index, "Category", lblCategory.Caption
    Reg.SaveSetting "Favourite " & Index, "From", cboFrom.ListIndex
    Reg.SaveSetting "Favourite " & Index, "To", cboTo.ListIndex
    
End Sub

Private Sub txtFrom_Change()
    CalculateAnswer
End Sub

' Calculate answer from txtFrom
Private Sub CalculateAnswer()
    Dim dblAns As Double    ' Define answer as a double
    
    ' Check if txtFrom is a number and combo's are not empty
    If IsNumeric(txtFrom.Text) And Not txtFrom.Text = "" And Not cboTo.Text = "" Then
        
        With Categorys(lblCategory.Caption)
            .strFromUnit = cboFrom.Text
            .strToUnit = cboTo.Text
            .dblValue = Val(txtFrom)
            dblAns = .Convert
            
            txtTo.Text = Round(dblAns, gbPrecision)
        End With
    End If
End Sub

