Attribute VB_Name = "MxCnt256"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxCnt256."
Public Const FFoCnt256$ = "Asc Chr Cnt"
Function Cnt256(S) As Long()
':Cnt256: :LngAy-with-256-ele
Dim O&(): ReDim O(255)
Dim J&: For J = 1 To Len(S)
    Dim A As Byte: A = Asc(Mid(S, J, 1))
    O(A) = O(A) + 1
Next
Cnt256 = O
End Function

Function DoCnt256(Cnt256&()) As Drs
DoCnt256 = Drs(FoCnt256, DyoCnt256(Cnt256))
End Function

Function FoCnt256() As String()
FoCnt256 = SyzSS(FFoCnt256)
End Function

Function DyoCnt256(Cnt256&()) As Variant()
Dim J%: For J = 0 To 255
    If Cnt256(J) > 0 Then
        PushI DyoCnt256, Array(J, Chr(J), Cnt256(J))
    End If
Next
End Function
