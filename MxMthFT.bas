Attribute VB_Name = "MxMthFT"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthFT."

Function MthFeiszSN(Src$(), Mthn) As Feis
Dim FmIx&, EIx&, Ix&, I
For Each I In Itr(MthIxyzN(Src, Mthn))
    Ix = I
    FmIx = Ix
   PushFei MthFeiszSN, Fei(FmIx, EndLix(Src, FmIx))
Next
End Function

Function MthFeiszMN(M As CodeModule, Mthn) As Feis
MthFeiszMN = MthFeiszSN(Src(M), Mthn)
End Function

Function MthFeis(Src$()) As Feis
Dim Ix&, FmIx&, EIx&, I
For Each I In MthIxItr(Src)
    Ix = I
    FmIx = Ix
    PushFei MthFeis, Fei(FmIx, EIx)
Next
End Function
