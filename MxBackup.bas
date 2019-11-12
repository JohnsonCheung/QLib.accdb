Attribute VB_Name = "MxBackup"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxBackup."

Sub BrwBkp()
BrwPth BkpzP(CPj)
End Sub

Function BackupFfn$(Ffn, Optional Msg$ = "Backup")
':XBUF: :X-Pfx #Back-Up-Ffn#
Const CSub$ = CMod & "BackupFfn"
Dim Tmpn$:       Tmpn = TmpNm
Dim TarFfn$:   TarFfn = BkFfn(Ffn)
Dim MsgFfn$:   MsgFfn = Pth(TarFfn) & "Msg.txt"
Dim MsgiFfn$: MsgiFfn = ParPthzFfn(TarFfn) & "MsgIdx.txt"
Dim Msgi$:       Msgi = "#" & Tmpn & vbTab & Msg
:                       CpyFfn Ffn, TarFfn
:                       WrtStr Msg, MsgFfn
:                       ApdStr Msgi, MsgiFfn
:                       BackupFfn = TarFfn
:                       InfLin CSub, "File is backuped", "As-file", TarFfn
End Function

Function BkPth$(Ffn)
':BkPth: :Pth #Backup-Path# ! Backup path of a Ffn
BkPth = AddFdrEns(BkHom(Ffn), TmpNm)
End Function

Sub BackupP(Optional Msg$ = "Backup")
':FunVerb-Backup: :FunVerb
':FunSfx-P: :FunSfx #CurPj#
BackupFfn Pjf(CPj), Msg
End Sub

Function BkFfn$(Ffn)
BkFfn = BkPth(Ffn) & Fn(Ffn)
End Function

Function BkHomP$()
BkHomP = BkHom(CPjf)
End Function

Function BkHom$(Ffn)
BkHom = EnsPth(Ffn & ".backup")
End Function

Function LasBkFfn$(Ffn)
Dim H$: H = BkHom(Ffn)
Dim F$(): F = FdrAyzIsInst(H)
Dim Fdr$: Fdr = MaxEle(F)
LasBkFfn = H & Fdr & "\" & Fn(Ffn)
End Function

Function BkFfnAy(Ffn) As String()
Dim H$: H = BkHom(Ffn)
Dim F$(): F = FdrAyzIsInst(H)
Dim Fn1$: Fn1 = Fn(Ffn)
Dim Fdr: For Each Fdr In Itr(F)
    Dim IFfn$: IFfn = H & Fdr & "\" & Fn1
    If HasFfn(IFfn) Then
        PushI BkFfnAy, IFfn
    End If
Next
End Function

Function BkpzP$(P As VBProject)
BkpzP = BkPth(Pjf(P))
End Function

Function BkRoot$(Pth)
BkRoot = AddFdr(Pth, ".Backup")
End Function

Function BkPjfAy() As String()
BkPjfAy = BkFfnAy(CPjf)
End Function

Function LasBkPjf$()
':FunAdj-Bk: :FunAdj #Backup#
':FunAdj-Las: :FunAdj #Last#
LasBkPjf = LasBkFfn(CPjf)
End Function
