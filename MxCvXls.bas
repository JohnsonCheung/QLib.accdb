Attribute VB_Name = "MxCvXls"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxCvXls."
Function CvLo(A) As ListObject
Set CvLo = A
End Function

Function CvRg(A) As Range
Set CvRg = A
End Function

Function CvWb(A) As Workbook
Set CvWb = A
End Function

Function CvWbs(A) As Workbooks
Set CvWbs = A
End Function

Function CvWs(A) As Worksheet
Set CvWs = A
End Function

