Attribute VB_Name = "MxRun"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRun."
Enum EmWaitRslt
    EiTimUp
    EiCnl
End Enum
Type WaitOpt
    TimOutSec As Integer
    ChkSec As Integer
    KeepFcmd As Boolean
End Type
Declare Function GetCurrentProcessId& Lib "Kernel32.dll" ()
'Declare Function GetProcessId& Lib "Kernel32.dll" (ProcessHandle&)
'Const Ps1Str$ = "function Get-ExcelProcessId { try { (Get-Process -Name Excel).Id } finally { @() } }" & vbCrLf & _
'"Stop-Process -Id (Get-ExcelProcessId)"

Function WaitOpt(TimOutSec%, ChkSec%, KeepFcmd As Boolean) As WaitOpt
With WaitOpt
.TimOutSec = TimOutSec
.ChkSec = ChkSec
.KeepFcmd = KeepFcmd
End With
End Function

Property Get DftWait() As WaitOpt
DftWait = WaitOpt(30, 5, False)
End Property

Sub KillProcessId(ProcessId&)
End Sub

Function RunFps1&(Fps1$, Optional PmStr$)
RunFps1 = RunFcmd("PowerShell", QteDbl(Fps1) & " " & PmStr)
End Function

Function RunFcmd&(Fcmd$, Optional PmStr$, Optional Sty As VbAppWinStyle = vbMaximizedFocus)
Dim Lin
    Lin = QteDbl(Fcmd) & PpdSpcIf(PmStr)
RunFcmd = Shell(Lin, Sty)
End Function

Function WaitFcmdw(Fcmdw$, W As WaitOpt, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus) As Boolean _
'Return True, if Fcmdw has generated the Fwaitg
Dim ProcessId&: ProcessId = Shell(Fcmdw, Sty)
Dim Fw$: Fw = Fwaitg(Fcmdw)
If WaitFwaitg(Fw, W.ChkSec, W.TimOutSec) Then
    Kill Fw
    WaitFcmdw = True
Else
    KillProcessId ProcessId
End If
If Not W.KeepFcmd Then Kill Fcmdw
End Function

Function WaitFwaitg(Fwaitg$, Optional ChkSec% = 10, Optional TimOutSec% = 60, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus) As Boolean _
'Return True, if Fwaitg is found.
Dim J%
For J = 1 To TimOutSec \ ChkSec
    If HasFfn(Fwaitg) Then
        Kill Fwaitg
        Exit Function
    End If
    If Not Wait(ChkSec%) Then Exit Function
Next
End Function

Sub Z_Fcmdw()
Debug.Print LineszFt(Fcmdw("Dir"))
End Sub
Function Fwaitg$(Fcmd$)
Fwaitg = Fcmd & ".wait.txt"
End Function

Function Fcmdw$(CmdLines$)
Dim T$: T = TmpFcmd
Dim EchoLin: EchoLin = FmtQQ("Echo > ""?""", Fwaitg(T))
Dim S$: S = CmdLines & vbCrLf & EchoLin
Fcmdw = WrtStr(S, T)
End Function

Sub Z_RunFcmd()
RunFcmd "Cmd"
MsgBox "AA"
End Sub

Function Wait(Optional Sec% = 1) As EmWaitRslt
Dim Till As Date: Till = AftSec(Sec)
Wait = IIf(Xls.Wait(Till), EiTimUp, EiCnl)
End Function

Function AftSec(Sec%) As Date 'Return the Date after Sec from Now
AftSec = DateAdd("S", Sec, Now)
End Function
Function Pipe(Pm, Mthnn$)
Dim O: Asg Pm, O
Dim I
For Each I In Ny(Mthnn)
   Asg Run(I, O), O
Next
Asg O, Pipe
End Function

Function RunAvzIgnEr(Mthn, Av())
If Si(Av) > 9 Then Thw CSub, "Si(Av) should be 0-9", "Si(Av)", Si(Av)
On Error Resume Next
RunAv Mthn, Av
End Function
Function RunAv(Mthn, Av())
Dim O
Select Case Si(Av)
Case 0: O = Run(Mthn)
Case 1: O = Run(Mthn, Av(0))
Case 2: O = Run(Mthn, Av(0), Av(1))
Case 3: O = Run(Mthn, Av(0), Av(1), Av(2))
Case 4: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case 9: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7), Av(8))
Case Else: Thw CSub, "UB-Av should be <= 8", "UB-Si Mthn", UB(Av), Mthn
End Select
RunAv = O
End Function


Sub RunCdLy(CdLy$())
RunCd JnCrLf(CdLy)
End Sub

Sub RunCd(CdLines$)
Dim N$: N = "Z_" & TmpNm
AddMthzCd N, CdLines
Run N
End Sub

Function RunCdMd() As CodeModule
'EnsMd "ZTmpModForRun"
End Function
Sub AddMthzCd(Mthn, CdLines$)
RunCdMd.AddFromString Mthl(Mthn, CdLines)
End Sub
Function Mthl$(Mthn, CdLines$)
Dim Lines$, L1$, L2$
L1 = "Sub Z_" & Mthn & "()"
L2 = "End Sub"
Mthl = L1 & vbCrLf & CdLines & vbCrLf & L2
End Function

Function Y_CdLines$()
Y_CdLines = "MsgBox Now"
End Function


Sub TimFun(FunNN)
Dim b!, E!, F
For Each F In TermAy(FunNN)
    b = Timer
    Run F
    E = Timer
    Debug.Print F, "<-- Run"; E - b
Next
End Sub

Sub Z_TimFun()
TimFun "ZZA ZZB"
End Sub

Sub ZZA()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        Debug.Print I
    Next
Next
End Sub
Sub ZZB()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        Debug.Print I
    Next
Next
End Sub
