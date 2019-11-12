Attribute VB_Name = "MxDiOp"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDiOp."
Sub PushDiS12o(D As Dictionary, A As S12Opt)
If A.Som Then PushDiS12 D, A.S12
End Sub

Sub PushDiS12(D As Dictionary, A As S12)
D.Add A.S1, A.S2
End Sub
