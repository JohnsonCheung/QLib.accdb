Attribute VB_Name = "Mx2AyOp"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "Mx2AyOp."
Function AyIntersect(A, B)
AyIntersect = ResiU(A)
If Si(A) = 0 Then Exit Function
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If HasEle(B, V) Then PushI AyIntersect, V
Next
End Function

Function SyMinus(A$(), B$()) As String()
SyMinus = AyMinus(A, B)
End Function

Function AyMinus(A, B)
If Si(B) = 0 Then AyMinus = A: Exit Function
AyMinus = ResiU(A)
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If Not HasEle(B, V) Then
        PushI AyMinus, V
    End If
Next
End Function

