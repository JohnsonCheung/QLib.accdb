Attribute VB_Name = "MxFmlCol"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxFmlCol."

Function AddFml(A As Drs, NewFld$, FunNm$, PmAy$()) As Drs
Dim Dy(): Dy = A.Dy
If Si(Dy) = 0 Then AddFml = A: Exit Function
Dim Dr, U&, Ixy1&(), Av()
Ixy1 = Ixy(A.Fny, PmAy)
U = UB(A.Fny)
For Each Dr In Dy
    If UB(Dr) <> U Then Thw CSub, "Dr-Si is diff", "Dr-Si U", UB(Dr), U
    Av = AwIxy(Dr, Ixy1)
    Push Dr, RunAv(FunNm, Av)
Next
AddFml = Drs(AddEleS(A.Fny, NewFld), Dy)
End Function

Function AddFmlSy(A As Drs, FmlSy$()) As Drs
Dim O As Drs: O = A
Dim NewFld$, FunNm$, PmAy$(), Fml$, I
For Each I In Itr(FmlSy)
    Fml = I
    NewFld = Bef(Fml, "=")
    FunNm = IsBet(Fml, "=", "(")
    PmAy = SplitComma(BetBkt(Fml))
    O = AddFml(O, NewFld, FunNm, PmAy)
Next
End Function
