Attribute VB_Name = "MxAcsOp"
Option Explicit
Option Compare Text
Const CLib$ = "QAcs."
Const CMod$ = CLib & "MxAcsOp."

Sub CpyFrm(A As Access.Application, Fb)
Dim I As AccessObject: For Each I In A.CodeProject.AllForms
    A.DoCmd.CopyObject Fb, , acForm, I.Name
Next
End Sub

Sub CpyMdzA(A As Access.Application, ToFb)
Dim I As AccessObject
For Each I In A.CodeProject.AllModules
    A.DoCmd.CopyObject ToFb, , acModule, I.Name
Next
End Sub

Sub CpyTzA(A As Access.Application, ToFb)
Dim I As AccessObject: For Each I In A.CodeData.AllTables
    A.DoCmd.CopyObject ToFb, , acTable, I.Name
Next
End Sub

Sub BrwFb(Fb)
Static Acs As New Access.Application
OpnFb Acs, Fb
Acs.Visible = True
End Sub

Sub BrwT(D As Database, T)
AcszDb(D).DoCmd.OpenTable T
End Sub

Sub BrwTT(D As Database, TT$)
Dim T
For Each T In ItrzTT(TT)
    BrwT D, T
Next
End Sub

Sub ClsAllTbl(D As Database)
Dim A As Access.Application: Set A = AcszDb(D)
Dim T: For Each T In A.CodeData.AllTables
    ClsTzA A, T
Next
End Sub

Sub ClsDbzAcs(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
End Sub

Sub ClsT(D As Database, T)
AcszDb(D).DoCmd.Close acTable, T
End Sub

Sub ClsTT(D As Database, TT$)
Dim A As Access.Application: Set A = AcszDb(D)
Dim T: For Each T In TermAy(TT)
    ClsTzA A, T
Next
End Sub

Sub ClsTzA(A As Access.Application, T)
A.DoCmd.Close acTable, T
End Sub


