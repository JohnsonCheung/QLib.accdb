Attribute VB_Name = "MxRfOp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxRfOp."

Sub CpyPjRfToPj(P As VBProject, ToPj As VBProject)
AddRfzAy ToPj, RffyzP(P)
End Sub

Sub RmvRfzRfnn(P As VBProject, Rfnn$)
Dim N
For Each N In Ny(Rfnn)
    P.References.Remove N
Next
SavPj P
End Sub

Sub RmvRf(P As VBProject, Rfn)
If HasRf(P, Rfn) Then P.References.Remove P.References(Rfn)
End Sub
