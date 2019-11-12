Attribute VB_Name = "MxLoNy"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxLoNy."

Function LoNyzWs(S As Excel.Worksheet) As String()
LoNyzWs = Itn(S.ListObjects)
End Function

Function LoNyzWb(B As Workbook) As String()
Dim S As Worksheet: For Each S In B.Sheets
    PushIAy LoNyzWb, LoNyzWs(S)
Next
End Function

Function FstWbLoNy() As String()
FstWbLoNy = LoNyzWb(FstWb)
End Function

