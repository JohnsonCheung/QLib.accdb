Attribute VB_Name = "MxGit"
Option Compare Text
Option Explicit
Const CLib$ = "QGit."
Const CMod$ = CLib & "MxGit."
Const Fgit$ = "C:\Program Files\Git\Cmd\git.Exe"
Const FgitQ$ = vbDblQ & Fgit & vbDblQ
Function Fcmd$(CmdPfx$, CmdStr$)
Fcmd = WrtStr(CmdStr, TmpFcmd(CmdPfx))
End Function
Function GitCmitCmdStr$(Optional Msg$ = "Commit", Optional ReInit As Boolean)
Dim X As New Bfr
    Dim P$:     P = SrcpP
    X.Var "Cd """ & P & """"
    If ReInit Then X "Rd .git /s/q"
    X.Var FmtQQ("? init", FgitQ)       'If already init, it will do nothing
    X.Var FmtQQ("? add -A", FgitQ)
    X.Var FmtQQ("? commit -m ""?""", FgitQ, Msg)
    X.Var "Pause"
GitCmitCmdStr = X.Lines
End Function
Sub GitCmit(Optional Msg$ = "commit", Optional ReInit As Boolean)
Dim CmdStr$: CmdStr = GitCmitCmdStr(Msg, ReInit)
ShellMax Fcmd("Cmit", CmdStr)
End Sub
Function GitPushCmdStr$()
Dim X As New Bfr
    Dim P$: P = SrcpP
    X.Var FmtQQ("Cd ""?""", P)
    X.Var FmtQQ("? push -u https://johnsoncheung@github.com/johnsoncheung/?.git master", FgitQ, PjnzSrcp(P))
    X.Var "Pause ....."
GitPushCmdStr = X.Lines
End Function

Sub GitPush()
ShellMax Fcmd("Push", GitPushCmdStr)
End Sub

Function HasInternet() As Boolean
Stop
End Function

Function PjnzSrcp$(Srcp$)
Dim P$: P = RmvPthSfx(Srcp)
If Ext(P) <> ".src" Then Thw CSub, "Not source path", "CmitgPth", Srcp
PjnzSrcp = RmvExt(Fn(P))
End Function
