VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppCaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

' ==== Save infos from "Application.Caller" : ====
' Classe SIMPLE créée en remplacement d'une variable de Type
' car les variable Type ne peuvent pas être passées à une Function ou Sub
Public CellsCount As Long
Public RowsCount As Long
Public ColsCount As Long
Public FileName As String
Public sheetName As String
' Full cells' addresses
Public Address As String
' Full Excel path format included the Sheet'name
Public xlPath As String
Public range As Variant
' The ALL formula
Public Formula As String
'Just the Windows file path
Public filePath As String
'========

Private Sub Class_Initialize()
  On Error GoTo ErrHandler

Exit Sub

ErrHandler:
  Debug.Print err.Number, err.Description
  err.Clear
  Resume Next
End Sub
