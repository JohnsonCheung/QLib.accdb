Attribute VB_Name = "MxLTD"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxLTD."
Public Const FFoLTD$ = "L T1 Dta"
Type TDoLTD: D As Drs: End Type
Function TDoLTDzInd(IndentSrc$()) As TDoLTD
TDoLTDzInd = TDoLTDzH(TDoLTDH(IndentSrc))
End Function
Function TDoLTD(Src$()) As TDoLTD
TDoLTD.D = DrszFF(FFoLTD, DyoLTD(Src))
End Function

Function TDoLTDzH(A As TDoLTDH) As TDoLTD
TDoLTDzH.D = DwFalseExl(A.D, "IsHdr")
End Function

Private Function DyoLTD(Src$()) As Variant()
'Ret:: Dy{L T1 Dta}
Dim L&, Dta$, T1$, Lin
For Each Lin In Itr(Src)
    L = L + 1
    If Fst2Chr(LTrim(L)) = "--" Then GoTo X
    T1 = T1zS(Lin)
    Dta = RmvT1(Lin)
    PushI DyoLTD, Array(L, T1, Dta)
X:
Next
End Function

