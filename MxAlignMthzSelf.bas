Attribute VB_Name = "MxAlignMthzSelf"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxAlignMthzSelf."

Sub AlignMthzSelf()
'Cpy Md
    Const TMdn$ = "QIde_B_AlignMth"      ' #The-Mdn
    Const TmMdn$ = "ATmp"                ' #Tmp-Mdn
:                                    EnsCls CPj, TmMdn
    Dim FmM As CodeModule: Set FmM = Md(TMdn)
    Dim ToM As CodeModule: Set ToM = Md(TmMdn)
    Dim OIsCpy As Boolean:  OIsCpy = CpyMd(FmM, ToM)
:                                    If OIsCpy Then MsgBox "Copied": Exit Sub

'Align
    Const TMthn$ = "AlignMthzLno"      ' #The-Mthn
    Dim M As CodeModule: Set M = Md(TMdn)
    Dim MthLno&:        MthLno = MthLnozMM(M, TMthn)
    'ATmp.AlignMthzLno M, MthLno, Upd:=EiUpdAndRpt, IsUpdSelf:=True
End Sub

Sub AlignMthEoLnk()
Dim M As CodeModule: Set M = Md("QDao_Lnk_EoLnk")
Dim L&:                  L = MthLnozMM(M, "EoLnk")
:                            AlignMthzLno M, L
End Sub
