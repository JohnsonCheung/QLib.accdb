Attribute VB_Name = "MxCurXls"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxCurXls."
Function CWs() As Worksheet
Set CWs = Xls.ActiveSheet
End Function

Property Get CWb() As Workbook
Set CWb = Xls.ActiveWorkbook
End Property

