VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCategorys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A collection class of the class clsConversion."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsConversion"
Attribute VB_Ext_KEY = "Member0" ,"clsConversion"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function loadCategoryData() As Boolean
    Dim Category As clsCategory, Unit As clsUnit
    Dim strCategory As String, strTemp As String
    
    On Error GoTo ErrHandler
    
    loadCategoryData = True
    
    Open gbDataPath & "index.dat" For Input As #1
        Do Until EOF(1)
            Input #1, strCategory
            If strCategory = "" Then
                Close #1
                loadCategoryData = False
                Exit Function
            End If
                
            Open gbDataPath & strCategory For Input As #2
                Set Category = New clsCategory
                
                With Category
                    Input #2, strTemp
                    .strName = strTemp
                    Do Until EOF(2)
                        Line Input #2, strTemp
                        Set Unit = New clsUnit
                        
                        Unit.loadFromString strTemp
                        .Units.Add Unit
                        
                        Set Unit = Nothing
                    Loop
                End With
                
                Add Category
                Set Category = Nothing
            Close #2
        Loop
    Close #1
    
    Exit Function
    
ErrHandler:
    Select Case Err.Number
    Case 457
        MsgBox "You have already loaded the data!", vbExclamation
    Case Else
        MsgBox "Error: " & CStr(Err.Number) & vbCrLf & _
                "Description: " & Err.Description
    End Select
    
    Err.Clear
End Function

Public Sub saveCategorysData()
Attribute saveCategorysData.VB_Description = "Saves all the conversion in this colection files and generates an index file."
    Dim Category As clsCategory
    Dim strCategorys As String
    
    'Write Category files.
    For Each Category In mCol
        With Category
            'Add file to Categorys list.
            strCategorys = strCategorys & .strName & ".dat" & vbCrLf
            
            'Write current data to new file.
            Open gbDataPath & .strName & ".dat" For Output As #1
                'Print #1, .strName
                Print #1, .getCategoryData
            Close #1
        End With
    Next Category
    
    'Write the index.dat file.
    Open gbDataPath & "index.dat" For Output As #1
        Print #1, strCategorys
    Close #1
End Sub


Public Function Add(Category As clsCategory) As clsCategory
    mCol.Add Category, Category.strName

    'return the object created
    Set Add = Category
End Function

Public Property Get Item(vntIndexKey As Variant) As clsCategory
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

