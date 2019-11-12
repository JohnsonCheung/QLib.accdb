Attribute VB_Name = "MxFmtMulLinSrc"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxFmtMulLinSrc."
Function DyoMulStmtSrc(MulStmtSrc$()) As Variant()
Dim ODy(): ODy = DyoLyWithColon(MulStmtSrc)
Dim Dr, J%, I&
For Each Dr In ODy
    For J = 0 To UB(Dr) - 1
        Dr(J) = Dr(J) & ":"
    Next
    ODy(I) = Dr
    I = I + 1
Next
DyoMulStmtSrc = ODy
End Function

Function FmtMulStmtSrc(MulStmtSrc$()) As String()
FmtMulStmtSrc = FmtDy(DyoMulStmtSrc(MulStmtSrc), MaxColWdt:=200)
End Function
