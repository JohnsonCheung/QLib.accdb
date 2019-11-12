Attribute VB_Name = "MxPcwsSrc"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPcwsSrc."
Sub Worksheet_SelectionChange(ByVal Target As Range)
PutPcwsChd Target
End Sub
    
