VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniConverter Editor"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Category"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&New"
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "D&elete"
         Height          =   315
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cboCategoryName 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Units"
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   4215
      Begin VB.TextBox txtOffset 
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtRelationship 
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox lstUnits 
         Height          =   1035
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Delete"
         Height          =   615
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   615
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   2295
         X2              =   2295
         Y1              =   245
         Y2              =   2045
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   2280
         X2              =   2280
         Y1              =   240
         Y2              =   2040
      End
      Begin VB.Label Label5 
         Caption         =   "&Offset:"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "&Relationship:"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "&Name:"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempCategorys As clsCategorys
Dim blnChanged As Boolean
Dim blnEditing As Boolean
Dim tmpUnit As clsUnit

Private Sub Form_Load()
    Set TempCategorys = New clsCategorys
    TempCategorys.loadCategoryData
    loadCategorys
    lockTextBoxes True
    
    'load pictures.
    cmdApply.Picture = LoadResPicture("Apply", vbResBitmap)
    cmdRemove.Picture = LoadResPicture("Delete", vbResBitmap)
    cmdEdit.Picture = LoadResPicture("Edit", vbResBitmap)
End Sub

Sub startEdit()
    lockTextBoxes False
    blnEditing = True
    
    cmdEdit.Caption = "Canc&el"
    cmdEdit.Picture = LoadResPicture("Cancel", vbResBitmap)
    cmdApply.Visible = True
    cmdRemove.Enabled = False
    lstUnits.Enabled = False
    cmdAdd.Enabled = False
    cboCategoryName.Enabled = False
    cmdOK.Enabled = False
    cmdOK.Default = False
    cmdApply.Default = True
    cmdCancel.Enabled = False
    cmdCancel.Cancel = True
    cmdEdit.Cancel = True
    
    txtName.SetFocus
End Sub

Sub stopEdit()
    blnEditing = False
    cmdEdit.Caption = "&Edit"
    cmdEdit.Picture = LoadResPicture("Edit", vbResBitmap)
    cmdEdit.Cancel = False
    
    cmdApply.Visible = False
    cmdApply.Default = False
    
    cmdRemove.Enabled = True
    lstUnits.Enabled = True
    cmdAdd.Enabled = True
    cboCategoryName.Enabled = True
    cmdOK.Enabled = True
    cmdOK.Default = True
    cmdCancel.Enabled = True
    cmdCancel.Cancel = True
    lockTextBoxes True
End Sub

Sub clearTextBoxes()
    txtName = ""
    txtOffset = ""
    txtRelationship = ""
End Sub

Sub loadCategorys()
    Dim Category As clsCategory
    
    cboCategoryName.Clear
    For Each Category In TempCategorys
        cboCategoryName.AddItem Category.strName
    Next Category
    
    cboCategoryName.ListIndex = IIf(cboCategoryName.ListCount = 0, "-1", 0)
    Set Category = Nothing
End Sub

Sub loadUnits()
    Dim Unit As clsUnit
    
    lstUnits.Clear
    lstUnits.AddItem "<new>"
    lstUnits.ItemData(0) = 1
    With TempCategorys(cboCategoryName.Text)
        For Each Unit In .Units
            lstUnits.AddItem Unit.strName
        Next Unit
    End With
    
    Set Unit = Nothing
End Sub

Sub loadTextBoxes()
    With TempCategorys(cboCategoryName.Text).Units(lstUnits.Text)
        txtName = .strName
        txtRelationship = .dblRelation
        txtOffset = .dblOffset
    End With
End Sub

Sub lockTextBoxes(blnLocked As Boolean)
    If blnLocked Then
        txtName.Locked = True
        txtName.BackColor = vbButtonFace
        txtName.ForeColor = &H808080
        txtOffset.Locked = True
        txtOffset.BackColor = vbButtonFace
        txtOffset.ForeColor = &H808080
        txtRelationship.Locked = True
        txtRelationship.BackColor = vbButtonFace
        txtRelationship.ForeColor = &H808080
    Else
        txtName.Locked = False
        txtName.BackColor = vbWindowBackground
        txtName.ForeColor = 0
        txtOffset.Locked = False
        txtOffset.BackColor = vbWindowBackground
        txtOffset.ForeColor = 0
        txtRelationship.Locked = False
        txtRelationship.BackColor = vbWindowBackground
        txtRelationship.ForeColor = 0
    End If
End Sub

Private Sub cboCategoryName_Click()
    If cboCategoryName.ListIndex >= 0 Then
        loadUnits
        clearTextBoxes
    End If
End Sub

Private Sub cmdAdd_Click()
    'Add new Converion type.
    Dim tmpCategory As clsCategory
    Dim strName As String
    
    strName = InputBox$("Enter a name for the new Category type:", APP_NAME)
    If strName = "" Then
        Exit Sub
    Else
        Set tmpCategory = New clsCategory
        tmpCategory.strName = strName
        
        TempCategorys.Add tmpCategory
        loadCategorys
        blnChanged = True
    End If
End Sub

Private Sub cmdApply_Click()
    Dim blnOK As Boolean
    
    blnOK = True
    
    If lstUnits.Text <> "<new>" Then
        'Copy tmpUnit to storage.
        With TempCategorys(cboCategoryName.Text).Units(lstUnits.Text)
            .strName = txtName
            .dblOffset = txtOffset
            .dblRelation = txtRelationship
        End With
        
        Set tmpUnit = Nothing
    Else
        ' Validate.
        If txtName.Text = "" Then
            MsgBox "Missing unit name.", vbExclamation, APP_NAME
            txtName.SetFocus
            blnOK = False
        ElseIf Val(txtRelationship.Text) = 0 Then
            MsgBox "Relationship cannot be zero.", vbExclamation, APP_NAME
            txtRelationship.SetFocus
            blnOK = False
        ElseIf Not IsNumeric(txtOffset.Text) Then
            MsgBox "Offset must be a number.", vbExclamation, APP_NAME
            txtOffset.SetFocus
            blnOK = False
        End If
        
        If blnOK Then
            'Add new unit.
            Set tmpUnit = New clsUnit
            With tmpUnit
                .strName = txtName
                .dblRelation = txtRelationship
                .dblOffset = Val(txtOffset)
            End With
                
            TempCategorys(cboCategoryName.Text).Units.Add tmpUnit
            
            Set tmpUnit = Nothing
        End If
    End If
    
    If blnOK Then
        stopEdit
        blnChanged = True
        'Add to list
        loadUnits
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If (MsgBox("Deleting will remove this Category and all units, continue?", vbYesNo + vbQuestion) = vbYes) And cboCategoryName.ListIndex >= 0 Then
        TempCategorys.Remove cboCategoryName.Text
        loadCategorys
    End If
End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        If lstUnits.ListIndex > 0 Then
            'Edit an item...
            
            'Copy existing data.
            Set tmpUnit = New clsUnit
            
            With tmpUnit
                .strName = txtName
                .dblRelation = Val(txtRelationship)
                .dblOffset = Val(txtOffset)
            End With
            
            startEdit
        End If
    Else
        'Cancel the edit.
        
        If lstUnits.Text <> "<new>" Then
            'Restore old data.
            With tmpUnit
                txtName = .strName
                txtRelationship = .dblRelation
                txtOffset = .dblOffset
            End With
            Set tmpUnit = Nothing
        End If
        
        stopEdit
        
    End If
End Sub

Private Sub cmdOK_Click()
    ' Check that all Categorys have at least 1 unit.
    Dim i As Integer
    Dim strConvName As String
    
    For i = 1 To TempCategorys.Count
        If TempCategorys(i).Units.Count = 0 Then
            strConvName = TempCategorys(i).strName
            Exit For
        End If
    Next i
    
    If strConvName = "" Then
        ' Save and close.
        TempCategorys.saveCategorysData
        Unload Me
    Else
        MsgBox "Category " & Chr$(34) & strConvName & Chr$(34) & " has no units defined, please define at least one!", vbExclamation, APP_NAME
    End If
End Sub

Private Sub cmdRemove_Click()
    If lstUnits.ListIndex > 0 Then
        TempCategorys(cboCategoryName.ListIndex + 1).Units.Remove lstUnits.ListIndex
        lstUnits.RemoveItem lstUnits.ListIndex
        lstUnits.ListIndex = -1
        txtName = "": txtRelationship = "": txtOffset = ""
        blnChanged = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEdit = Nothing
End Sub

Private Sub lstUnits_Click()
    If lstUnits.ListIndex > 0 Then
        loadTextBoxes
    Else
        'Add a new unit...
        'Clear the unit textboxes..
        clearTextBoxes
        
        startEdit
    End If
End Sub

