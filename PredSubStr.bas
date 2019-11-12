VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredSubStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Implements IPred
Const CLib$ = "QVb."
Const CMod$ = CLib & "PredSubStr."
Private A$

Sub Init(SubStr)
A = SubStr
End Sub

Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = HasSubStr(V, A)
End Function
