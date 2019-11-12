VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredLikAy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IPred
Const CLib$ = "QVb."
Const CMod$ = CLib & "PredLikAy."
Private A$(), Emp As Boolean
Sub Init(LikAy$())
A = LikAy
Emp = Si(A) = 0
End Sub

Function Pred(V) As Boolean
If Emp Then Pred = True: Exit Function
Dim Lik
For Each Lik In A
    If V Like Lik Then Pred = True: Exit Function
Next
End Function

Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = Pred(V)
End Function
