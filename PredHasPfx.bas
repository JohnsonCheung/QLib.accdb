VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredHasPfx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IPred
Const CLib$ = "QVb."
Const CMod$ = CLib & "PredHasPfx."
Private P$
Friend Sub Init(Pfx$)
P = Pfx
End Sub
Function Pred(V) As Boolean
Pred = IPred_Pred(V)
End Function
Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = HasPfx(V, P)
End Function
