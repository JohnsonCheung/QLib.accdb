Attribute VB_Name = "MxXlsVis"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsVis."
Private A() As Boolean
Sub PushXlsVis()
PushI A, Xls.Visible
End Sub
Sub PopXlsVis()
Xls.Visible = PopI(A)
End Sub
Sub PushXlsVisHid()
PushXlsVis
Xls.Visible = False
End Sub

Function ShwWb(A As Workbook) As Workbook
ShwXls A.Application
Set ShwWb = A
End Function

Function ShwXls(A As Excel.Application) As Excel.Application
If Not A.Visible Then A.Visible = True
Set ShwXls = A
End Function

Function ShwRg(A As Range) As Range
ShwXls A.Application
Set ShwRg = A
End Function

Function ShwLo(A As ListObject) As ListObject
ShwXls A.Application
Set ShwLo = A
End Function

Function ShwWs(A As Worksheet) As Worksheet
ShwXls A.Application
Set ShwWs = A
End Function


