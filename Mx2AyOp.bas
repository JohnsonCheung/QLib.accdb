Attribute VB_Name = "Mx2AyOp"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "Mx2AyOp."
Function AyIntersect(A, b)
AyIntersect = ResiU(A)
If Si(A) = 0 Then Exit Function
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If HasEle(b, V) Then PushI AyIntersect, V
Next
End Function

Function SyMinus(A$(), b$()) As String()
SyMinus = AyMinus(A, b)
End Function

Function AyMinus(A, b)
If Si(b) = 0 Then AyMinus = A: Exit Function
AyMinus = ResiU(A)
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If Not HasEle(b, V) Then
        PushI AyMinus, V
    End If
Next
End Function

