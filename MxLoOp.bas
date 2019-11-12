Attribute VB_Name = "MxLoOp"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxLoOp."

Sub BrwLo(A As ListObject)
BrwDrs DrszLo(A)
End Sub

Sub InsLcBef(L As ListObject, C$, BefCol$)
Dim Cno%: Cno = CnozLc(L, C)
EntColRgzLc(L, C).Insert
Lc(L, C).Name = C
End Sub

Sub InsLcAft(L As ListObject, C$, AftCol$)
'Do #Ins-ListObjectCol-AftCol#
End Sub

Sub KeepFstLc(L As ListObject)
Dim J%
For J = L.ListColumns.Count To 2 Step -1
    L.ListColumns(J).Delete
Next
End Sub

Sub KeepFstLr(A As ListObject)
Dim J%
For J = A.ListRows.Count To 2 Step -1
    A.ListRows(J).Delete
Next
End Sub


Function LoPc(A As ListObject) As PivotCache
Dim O As PivotCache
Set O = WbzLo(A).PivotCaches.Create(xlDatabase, A.Name, 6)
O.MissingItemsLimit = xlMissingItemsNone
Set LoPc = O
End Function

Function LoQt(A As ListObject) As QueryTable
On Error Resume Next
Set LoQt = A.QueryTable
End Function

Function R1Lo&(A As ListObject, Optional InclHdr As Boolean)
If IsLozNoDta(A) Then
   R1Lo = A.ListColumns(1).Range.Row + 1
   Exit Function
End If
R1Lo = A.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function R2Lo&(A As ListObject, Optional InclTot As Boolean)
If IsLozNoDta(A) Then
   R2Lo = R1Lo(A)
   Exit Function
End If
R2Lo = A.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Function SqzLo(A As ListObject) As Variant()
If IsNothing(A.DataBodyRange) Then Exit Function
SqzLo = A.DataBodyRange.Value
End Function

Function WszLo(A As ListObject) As Worksheet
Set WszLo = A.Parent
End Function

Function WbzLo(A As ListObject) As Workbook
Set WbzLo = WbzWs(WszLo(A))
End Function

Function WsCnozLc&(A As ListObject, Col)
WsCnozLc = A.ListColumns(Col).Range.Column
End Function

Function LonzT$(T)
LonzT = "T_" & RmvFstNonLetter(T)
End Function

Sub Z_KeepFc()
KeepFstLc ShwLo(SampLo)
End Sub

Sub Z_AutoFitLo()
Dim Ws As Worksheet: Set Ws = NewWs
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "A"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "A")
Ws.Range("A1:B2").Value = Sq
AutoFit LozWsDta(Ws)
ClsWsNoSav Ws
End Sub

Sub Z_BrwLo()
BrwLo SampLo
Stop
End Sub

Sub Z_PtzLo()
Dim At As Range, Lo As ListObject
Set Lo = SampLo
ShwPt PtzLo(Lo, At, "A B", "C D", "F", "E")
Stop
End Sub

