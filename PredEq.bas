VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredEq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IPred
Const CLib$ = "QVb."
Const CMod$ = CLib & "PredEq."
Private A
Sub Ini(Eq)
A = Eq
End Sub

Function Pred(V) As Boolean
Pred = V = A
End Function

Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = Pred(V)
End Function


