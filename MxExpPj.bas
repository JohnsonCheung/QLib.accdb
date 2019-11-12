Attribute VB_Name = "MxExpPj"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxExpPj."

Sub ExpRf(P As VBProject)
WrtAy RfSrc(P), Frf(P)
End Sub


Sub ExpPjf(Pjf, Optional Xls As Excel.Application, Optional Acs As Access.Application)
Stamp "ExpPj: Begin"
Stamp "ExpPj: Pjf " & Pjf
Select Case True
Case IsFxa(Pjf): ExpFxa Pjf
Case IsFba(Pjf): ExpFba Pjf, Acs
End Select
Stamp "ExpPj: End"
End Sub

Sub ExpFba(Fba, Optional Acs As Access.Application)
CpyFfnzToPth Fba, EnsPth(SrcpzPjf(Fba))
Dim A As Access.Application: Set A = DftAcs(Acs)
OpnFb A, Fba
Dim Pj As VBProject: Set Pj = A.Vbe.ActiveVBProject
ExpPjzP Pj
QuitAcs A
End Sub

Sub ExpFxa(Fxa)
ExpPjzP PjzFxa(Fxa)
End Sub

Sub ExpPjP()
ExpPjzP CPj
End Sub

Sub ExpSrc(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    DoEvents
    C.Export SrcFfn(C)
Next
End Sub

Sub ExpFrm(P As VBProject)
If Not IsFbaPj(P) Then Exit Sub
Stop
End Sub

Sub ExpFrmzAcs(A As Access.Application, ToPth$)
Dim F As AccessObject
For Each F In A.CurrentProject.AllForms
    A.SaveAsText acForm, F.Name, ToPth & F.Name & ".frm.txt"
Next
End Sub

Sub ExpPjzP(Pj As VBProject)
Dim P$: P = SrcpzP(Pj)
InfLin CSub, "... Clr src pth":       EnsPthAll P
                                      ClrPthFil P
InfLin CSub, "... Cpy pj to src pth": CpyFfnzToPth Pj.FileName, P
InfLin CSub, "... Exp src":           ExpSrc Pj
InfLin CSub, "... Exp rf":            ExpRf Pj
InfLin CSub, "... Exp frm":           ExpFrm Pj
InfLin CSub, "Done"
End Sub

