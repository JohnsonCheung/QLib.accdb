Attribute VB_Name = "MxCntg"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxCntg."

Function CntDyWhGt1zAy(Ay) As Variant()
CntDyWhGt1zAy = CntDyWhGt1(DyzDi(DiKqCnt(Ay)))
End Function

Function DyoCntg(Ay) As Variant()
DyoCntg = DyzDi(DiKqCnt(Ay))
End Function

Sub Z_DyoCntg()
Dim A$(): A = SplitSpc("a a a b c b")
Dim Act(): Act = DyoCntg(A)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
GoSub Tst
Exit Sub
Tst:
    Act = DyoCntg(A)
    Ass IsEqAy(Act, Ept)
    Return
End Sub

Function SumSi&(Ay)
Dim I, O&
For Each I In Itr(Ay)
    O = O + Len(I)
Next
SumSi = O
End Function
Sub Z_CntSiLin()
Debug.Print CntSiLin(SrczP(CPj))
End Sub
Function CntSiLin(Ay)
CntSiLin = "AyCntSi(" & Si(Ay) & "." & SumSi(Ay) & ")"
End Function
