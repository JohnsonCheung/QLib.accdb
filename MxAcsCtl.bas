Attribute VB_Name = "MxAcsCtl"
Option Explicit
Option Compare Text
Const CLib$ = "QAcs."
Const CMod$ = CLib & "MxAcsCtl."

Sub SetNoTabStop(A As Access.Form)
ForItrFun A.Controls, "CmdTurnOffTabStop"
End Sub

Function CvAcsCtl(A) As Access.Control
Set CvAcsCtl = A
End Function

Function CvAcsBtn(A) As Access.CommandButton
Set CvAcsBtn = A
End Function

Function CvAcsTgl(A) As Access.ToggleButton
Set CvAcsTgl = A
End Function


Sub TurnOffTabStop(A As Access.Control)
If Not HasPfx(A.Name, "Cmd") Then Exit Sub
Select Case True
Case IsBtn(A): CvBtn(A).TabStop = False
Case IsTglBtn(A): CvAcsTgl(A).TabStop = False
End Select
End Sub
