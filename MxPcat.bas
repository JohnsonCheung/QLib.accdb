Attribute VB_Name = "MxPcat"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPcat."
':Pcat: :Cml #Par-chd-At# ! It is a feature to allow a Drs to be shown in a ws
Sub Put_ChdDrs_ByTar_FmPcatBfr(Target As Range)
Dim A As Pcat, Fnd As Boolean: A = F_Pcat_ByTar_FmPcatBfr(Target, Fnd): If Not Fnd Then Exit Sub
Dim CurChdDrs As Drs
    Dim KeyDr():       KeyDr = DrzLoCell(A.UKeyLo, Target)
    Dim CurChdDy(): CurChdDy = DywKeyDr(A.ChdDy, KeyDr, A.KeyDy)
                   CurChdDrs = Drs(A.ChdFny, CurChdDy)
:                              PutDrsToLo CurChdDrs, A.ChdLo '<===
End Sub

