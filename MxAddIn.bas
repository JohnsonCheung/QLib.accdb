Attribute VB_Name = "MxAddIn"
Option Compare Text
Option Explicit
Const CNs$ = "AddIn"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxAddIn."

Function DoAddin(A As Excel.Application) As Drs
DoAddin = DrszItrPrpcc(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function

Sub DmpAddinX()
DmpAddinzX Xls
End Sub

Sub DmpAddinzX(X As Excel.Application)
DmpDrs DoAddin(X)
End Sub

Function WsAddin(X As Excel.Application) As Worksheet
Set WsAddin = ShwWs(WszDrs(DoAddin(X)))
End Function

Function Addin(X As Excel.Application, FxaNm) As Excel.Addin
Dim I As Excel.Addin
For Each I In X.AddIns
    If I.Name = FxaNm & ".xlam" Then Set Addin = I
Next
End Function
