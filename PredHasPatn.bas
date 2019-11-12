VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredHasPatn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IPred
Const CLib$ = "QVb."
Const CMod$ = CLib & "PredHasPatn."
Private Re As RegExp, AndRe1 As RegExp, AndRe2 As RegExp, HasAnd1 As Boolean, HasAnd2 As Boolean

Friend Sub Init(Patn$, Optional AndPatn1$, Optional AndPatn2$)
Set Re = Rx(Patn)
If AndPatn1 <> "" Then
    Set AndRe1 = Rx(AndPatn1)
    HasAnd1 = True
End If
If AndPatn2 <> "" Then
    Set AndRe2 = Rx(AndPatn2)
    HasAnd2 = True
End If
End Sub

Function Pred(V) As Boolean
Pred = IPred_Pred(V)
End Function

Function IPred_Pred(V As Variant) As Boolean
If Re.Test(V) Then
    If HasAnd1 Then
        If Not AndRe1.Test(V) Then Exit Function
    End If
    If HasAnd2 Then
        If Not AndRe2.Test(V) Then Exit Function
    End If
    IPred_Pred = True
End If
End Function
