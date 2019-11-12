Attribute VB_Name = "MxWbPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxWbPrp."
Function LasWs(A As Workbook) As Worksheet
Set LasWs = A.Sheets(A.Sheets.Count)
End Function


Function FstWs(A As Workbook) As Worksheet
Set FstWs = A.Sheets(1)
End Function


Function FxzWb$(A As Workbook)
Dim F$
F = A.FullName
If F = A.Name Then Exit Function
FxzWb = F
End Function



Function MainLo(A As Workbook) As ListObject
Dim O As Worksheet, Lo As ListObject
Set O = MainWs(A):              If IsNothing(O) Then Exit Function
Set MainLo = LozWs(O, "T_Main")
End Function

Function MainQt(A As Workbook) As QueryTable
Dim Lo As ListObject
Set Lo = MainLo(A): If IsNothing(A) Then Exit Function
Set MainQt = Lo.QueryTable
End Function


Function WszCdNm(A As Workbook, WsCdNm$) As Worksheet
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws.CodeName = WsCdNm Then Set WszCdNm = Ws: Exit Function
Next
End Function

Function WsCdNy(A As Workbook) As String()
WsCdNy = SyzItrPrp(A.Sheets, "CodeName")
End Function


Function WbFullNm$(A As Workbook)
On Error Resume Next
WbFullNm = A.FullName
End Function

Function PtNy(A As Workbook) As String()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushIAy PtNy, PtNyzWs(Ws)
Next
End Function


Function WszWb(A As Workbook, WsIx) As Worksheet
Set WszWb = A.Sheets(WsIx)
End Function

