Attribute VB_Name = "MxEle"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxEle."
Function FstEle(Ay)
If Si(Ay) = 0 Then Exit Function
Asg Ay(0), FstEle
End Function

Function LasEle(Ay)
Dim N&: N = Si(Ay)
If N = 0 Then
    Thw CSub, "No ele in Ay"
Else
    Asg Ay(N - 1), LasEle
End If
End Function

Function MinEle(Ay)
Dim O: O = FstEle(Ay)
Dim I: For Each I In Itr(Ay)
    If I < O Then O = I
Next
MinEle = O
End Function

Function MaxEle(Ay)
Dim O: O = FstEle(Ay)
Dim I: For Each I In Itr(Ay)
    If I > O Then O = I
Next
MaxEle = O
End Function

