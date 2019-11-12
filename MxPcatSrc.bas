Attribute VB_Name = "MxPcatSrc"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPcatSrc."
Sub Worksheet_SelectionChange(ByVal Target As Range)
Put_ChdDrs_ByTar_FmPcatBfr Target
End Sub

