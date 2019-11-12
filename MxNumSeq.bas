Attribute VB_Name = "MxNumSeq"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxNumSeq."

Function IntSeq(F%, T%) As Integer()
IntSeq = IntozFT(EmpIntAy, F, T)
End Function

Function IntozFT(Into, F, T)
Dim O: O = Into: ReDim O(Abs(T - F))
Dim S: S = IIf(T > F, 1, -1) ' Step
Dim V, I&: For V = F To T Step S
    O(I) = V
    I = I + 1
Next
IntozFT = O
End Function

Function LngSeq(F&, T&) As Long()
LngSeq = IntozFT(EmpLngAy, F, T)
End Function

Function IntSeqzN(N&, Optional Fm% = 0) As Integer()
Dim O%(): ReDim O(N - 1)
Dim J&
    For J = 0 To N - 1
        O(J) = J + Fm
    Next
IntSeqzN = O
End Function



Function IntSeqzU(U%) As Integer()
IntSeqzU = IntSeq(0, U)
End Function

Function LngSeqzU(U&) As Long()
LngSeqzU = LngSeq(0, U)
End Function
