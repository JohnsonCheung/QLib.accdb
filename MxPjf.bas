Attribute VB_Name = "MxPjf"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxPjf."
Public PjfAcs As New Access.Application
Public PjfXls As New Excel.Application

Sub ClsPjf(Pjf)
Const CSub$ = CMod & "ClsPjf"
Select Case True
Case IsFxa(Pjf): RmvPjzXlsPjf PjfXls, Pjf
Case IsFb(Pjf):  ClsDbzAcs PjfAcs
Case Else: Thw CSub, "Invalid Pjf, should be Fxa or Fb", "Pjf", Pjf
End Select
End Sub
Function VbezFba(Fba) As Vbe
'OpnPjf Pjf

End Function
Function VbezFxa(Fxa) As Vbe
'OpnPjf Pjf

End Function
Function VbezPjf(Pjf) As Vbe
Const CSub$ = CMod & "VbezPjf"
Select Case True
Case IsFxa(Pjf): Set VbezPjf = VbezFba(Pjf)
Case IsFba(Pjf):  Set VbezPjf = VbezFxa(Pjf)
Case Else: Thw CSub, "Invalid Pjf, should be Fxa or Fba", "Pjf", Pjf
End Select
End Function
Sub OpnPjf(Pjf)  ' Return either Xls.Application (Xls) or Acs.Application (Function-static)
Select Case True
Case IsFxa(Pjf): PjfXls.Workbooks.Open Pjf
Case IsFba(Pjf):  OpnFb PjfAcs, Pjf
Case Else: Stop
End Select
End Sub

Sub RmvPjzXlsPjf(Xls As Excel.Application, Pjf)
Dim Pj As VBProject
Set Pj = PjzPjf(Xls.Vbe, Pjf)
Pj.Collection.Remove Pj
End Sub

Function TmpFxa$(Optional Fdr$, Optional Fnn$)
TmpFxa = TmpFfn(".xlam", Fdr, Fnn)
End Function

