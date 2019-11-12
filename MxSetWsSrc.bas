Attribute VB_Name = "MxSetWsSrc"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxSetWsSrc."
Sub Z_AddWsSrczM()
Dim Wb As Workbook: Set Wb = Xls.Workbooks("Book1")
Dim Ws As Worksheet: Set Ws = Wb.Sheets("Sheet1")
AddWsSrczMdn Ws, "ShwDrsAtSheetSrc"
End Sub

Sub AddWsSrc(S As Worksheet, Srcl$)
RplMd MdzWs(S), Srcl
End Sub

Sub AddWsSrczM(S As Worksheet, SrcMd As CodeModule)
AddWsSrc S, Srcl(SrcMd)
End Sub

Sub AddWsSrczMdn(S As Worksheet, Mdn$)
AddWsSrczM S, Md(Mdn)
End Sub
