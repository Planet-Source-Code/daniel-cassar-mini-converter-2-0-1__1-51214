Attribute VB_Name = "modGlobal"
Option Explicit

Public Const APP_NAME = "MiniConverter"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public gbPrecision As Integer   ' Precision value
Public gbDataPath As String     ' Path to category data files
