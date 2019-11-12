Attribute VB_Name = "MxIdeTool"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeTool."

Function Spec() As String()
Erase XX
X "Bars"
X " AA A1 A2 A3"
X " BB B1 B2 B3"
X "Btns"
X " A1"
Spec = XX  '*Spec
Erase XX
End Function

Function BtnSpec() As String()
BtnSpec = IndentedLy(Spec, "Bars")
End Function

Sub InstallIdeTools()
EnsBtns BtnSpec
EnsClsLines "IdeTool", ToolClsCd
End Sub

Function ToolClsCd$()
Stop
End Function

Function ToolBarNy() As String()
ToolBarNy = AmT1(BtnSpec)
End Function

Sub RmvIdeTools()
RmvBarByNy ToolBarNy
End Sub

