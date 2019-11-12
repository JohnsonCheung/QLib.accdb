Attribute VB_Name = "MxFilOp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxFilOp."
Sub RplFfn(Ffn, ByFfn$)
BackupFfn Ffn
If DltFfnDone(Ffn) Then
    Name Ffn As ByFfn
End If
End Sub
Sub CpyPthzClr(FmPth$, ToPth$)
ThwIf_NoPth ToPth, CSub
ClrPthFil ToPth
Dim Ffn$, I
For Each I In FfnAy(FmPth)
    Ffn = I
    CpyFfnzToPth Ffn, ToPth
Next
End Sub

Sub CpyFfnUp(Ffn)
CpyFfnzToPth Ffn, ParPth(Ffn)
End Sub

Sub CpyFfnAyzToNxt(FfnAy$())
Dim I, Ffn$
For Each I In Itr(FfnAy)
    Ffn = I
    CpyFfnzToNxt Ffn
Next
End Sub

Function CpyFfnzToNxt$(Ffn)
Dim O$
O = NxtFfnzAva(Ffn)
CpyFfn Ffn, O
CpyFfnzToNxt = O
End Function

Sub CpyFfnzToPthIfDif(Ffn, ToPth$, Optional b As EmFilCmp = EmFilCmp.EiCmpEq)
Dim Fn$: Fn = FnzFfn(Ffn)
Dim ToFfn$: ToFfn = FfnzPthFn(ToPth, Fn)
If IsEqFfn(Ffn, ToFfn, b) Then Exit Sub
CpyFfnzToPth Ffn, ToPth, OvrWrt:=True
End Sub

Sub CpyFfnAyzIfDif(FfnAy$(), ToPth$, Optional b As EmFilCmp = EmFilCmp.EiCmpEq)
Dim I
For Each I In FfnAy
    CpyFfnzIfDif CStr(I), ToPth, b
Next
End Sub

Sub CpyFfn(Ffn, ToFfn$, Optional OvrWrt As Boolean)
Fso.GetFile(Ffn).Copy ToFfn, OvrWrt
End Sub

Function CpyFfnAy$(FfnAy$(), ToPth$, Optional OvrWrt As Boolean)
Dim Ffn$, I, P$, O$
P = EnsPthSfx(ToPth)
For Each I In FfnAy
    O = P & Fn(Ffn)
    CpyFfn Ffn, O, OvrWrt
Next
End Function
Function CpyFfnzToPth$(Ffn, ToPth$, Optional OvrWrt As Boolean)
CpyFfn Ffn, FfnzPthFn(ToPth, Fn(Ffn)), OvrWrt
End Function

Sub CpyFfnzIfDif(Ffn, ToFfn$, Optional b As EmFilCmp)
If IsEqFfn(Ffn, ToFfn, b) Then
    Dim M$: M = FmtQQ("? file", IIf(b = EiCmpEq, "Eq", "Same"))
    D LyzFunMsgNap(CSub, M, "FmFfn ToFfn", Ffn, ToFfn)
    Exit Sub
End If
CpyFfn Ffn, ToFfn, OvrWrt:=True
D LyzFunMsgNap(CSub, "File copied", "FmFfn ToFfn", Ffn, ToFfn)
End Sub

Sub DltFfnAyAyIf(FfnAy$())
Dim Ffn
For Each Ffn In Itr(FfnAy)
    DltFfnIf CStr(Ffn)
Next
End Sub

Sub DltFfn(Ffn)
On Error GoTo X
Kill Ffn
Exit Sub
X:
Thw CSub, "Cannot kill", "Ffn Er", Ffn, Err.Description
End Sub

Sub DltFfnIf(Ffn)
If HasFfn(Ffn) Then DltFfn Ffn
End Sub

Function DltFfnIfPrompt(Ffn, Msg$) As Boolean 'Return true if error
If NoFfn(Ffn) Then Exit Function
On Error GoTo X
Kill Ffn
Exit Function
X:
MsgBox "File [" & Ffn & "] cannot be deleted, " & vbCrLf & Msg
DltFfnIfPrompt = True
End Function

Function DltFfnDone(Ffn) As Boolean
On Error GoTo X
Kill Ffn
DltFfnDone = True
Exit Function
X:
End Function


Sub MovFilUp(Pth)
Dim I, Tar$
Tar$ = ParPth(Pth)
For Each I In Itr(FnAy(Pth))
    MovFfn CStr(I), Tar
Next
End Sub


Sub MovFfn(Ffn, ToPth$)
Fso.MoveFile Ffn, ToPth
End Sub
