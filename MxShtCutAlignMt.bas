Attribute VB_Name = "MxShtCutAlignMt"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxShtCutAlignMt."

Sub ACmdApply()
AlignMthzNm "QXls_Cmd_ApplyFilter", "CmdApply"
End Sub

Sub AU()
AlignMth Upd:=EiUpdAndRpt
End Sub

Sub AUO()
AlignMth Upd:=EiUpdOnly
End Sub
