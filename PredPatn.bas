VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredPatn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IPred
Const CLib$ = "QVb."
Const CMod$ = CLib & "PredPatn."
Private A As New RegExp
Friend Sub Init(Patn$)
A.Pattern = Patn
End Sub
Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = A.Test(V)
End Function
