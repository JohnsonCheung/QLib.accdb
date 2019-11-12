Attribute VB_Name = "MxPthOp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxPthOp."

Sub VcPth(Pth)
If NoPth(Pth) Then Exit Sub
Shell FmtQQ("Code.cmd ""?""", Pth), vbMaximizedFocus
End Sub

Sub BrwPth(Pth)
If NoPth(Pth) Then Exit Sub
ShellMax FmtQQ("Explorer ""?""", Pth)
End Sub


Sub ClrPth(Pth)
DltFfnAyAyIf FfnAy(Pth)
End Sub

Sub Z_ClrPthFil()
ClrPthFil TmpRoot
End Sub

Sub ClrPthFil(Pth)
If NoPth(Pth) Then Exit Sub
Dim F
For Each F In Itr(FfnAy(Pth))
   DltFfn F
Next
End Sub


Sub DltEmpPthR(Pth)
Dim Ay$(), I, J%
Lp:
    J = J + 1: If J > 10000 Then Stop
    Dim Dlt As Boolean: Dlt = False
    For Each I In Itr(EmpPthAyR(Pth))
        DltPthNoEr I
        Dlt = True
    Next
    If Dlt Then GoTo Lp
End Sub
Sub DltPthNoEr(Pth)
On Error Resume Next
RmDir Pth
End Sub

Sub DltEmpSubDir(Pth)
Dim SubPth
For Each SubPth In Itr(SubPthAy(Pth))
   DltPthIfEmp SubPth
Next
End Sub

Sub DltPthIfEmp(Pth)
If IsEmpPth(Pth) Then Exit Sub
RmDir Pth
End Sub

Sub RenPthAddFdrPfx(Pth, Pfx)
RenPth Pth, AddFdrPfx(Pth, Pfx)
End Sub

Sub RenPth(Pth, NewPth)
Fso.GetFolder(Pth).Name = NewPth
End Sub

Sub Z_DltEmpSubDir()
DltEmpSubDir TmpPth
End Sub
