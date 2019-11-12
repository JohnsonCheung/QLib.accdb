Attribute VB_Name = "MxPjDte"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxPjDte."

Function PjDtezFb(Fb, Optional Acs As Application) As Date
Dim A As Access.Application: Set A = DftAcs(Acs)
PjDtezFb = PjDtezAcs(A)
If IsNothing(Acs) Then QuitAcs A
End Function

Function PjDtezPjf(Pjf) As Date
Select Case True
Case IsFxa(Pjf): PjDtezPjf = DtezFfn(Pjf)
Case IsFb(Pjf): PjDtezPjf = PjDtezFb(Pjf)
Case Else: Stop
End Select
End Function

Function PjDtezAcs(A As Access.Application)
Dim O As Date
Dim M As Date
M = MaxzItrPrp(A.CurrentProject.AllForms, "DateModified")
O = Max(O, M)
O = Max(O, MaxzItrPrp(A.CurrentProject.AllModules, "DateModified"))
O = Max(O, MaxzItrPrp(A.CurrentProject.AllReports, "DateModified"))
PjDtezAcs = O
End Function
