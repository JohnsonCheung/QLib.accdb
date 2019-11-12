VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredInT1Sy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IPred
Const CLib$ = "QVb."
Const CMod$ = CLib & "PredInT1Sy."
Private A_AmT1$()
Friend Sub Init(AmT1$())
A_AmT1 = AmT1
End Sub

Function IPred_Pred(V As Variant) As Boolean
Dim I, Lin, T1$
Lin = V
For Each I In A_AmT1
    T1 = I
    If HasT1(Lin, T1) Then IPred_Pred = True: Exit Function
Next
End Function
