Attribute VB_Name = "MxDicab"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDicab."
Type DicAB
    A As Dictionary
    B As Dictionary
End Type
Function DicAB(A As Dictionary, B As Dictionary) As DicAB
ThwIf_Nothing A, "DicA", CSub
ThwIf_Nothing B, "DicB", CSub
With DicAB
    Set .A = A
    Set .B = B
End With
End Function
Function DicabzInKy(D As Dictionary, InKy) As DicAB
Dim K, A As New Dictionary, B As New Dictionary
For Each K In D.Keys
    If HasEle(InKy, K) Then
        A.Add K, D(K)
    Else
        B.Add K, D(K)
    End If
Next
DicabzInKy = DicAB(A, B)
End Function
