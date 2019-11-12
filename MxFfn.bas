Attribute VB_Name = "MxFfn"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxFfn."
Public Const FbExt$ = ".accdb"
Public Const FbExt1$ = ".mdb"
Public Const FbaExt$ = ".accdb"
Public Const FxaExt$ = ".xlam"
Enum EmFilCmp
    EiCmpEq
    EiCmpSam
End Enum
Public Const PthSep$ = "\"
Function CutPth$(Ffn)
Dim P%: P = InStrRev(Ffn, PthSep)
If P = 0 Then CutPth = Ffn: Exit Function
CutPth = Mid(Ffn, P + 1)
End Function
Function FnzFfn$(Ffn)
FnzFfn = CutPth(Ffn)
End Function

Function Fn$(Ffn)
Fn = CutPth(Ffn)
End Function

Function FfnUp$(Ffn)
FfnUp = ParPth(Pth(Ffn)) & Fn(Ffn)
End Function

Function Fnn$(Ffn)
Fnn = RmvExt(Fn(Ffn))
End Function

Function RmvExt$(Ffn)
Dim B$, C$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
RmvExt = Pth(Ffn) & C
End Function
Function IsExtInAp(Ffn, ParamArray Ap()) As Boolean

End Function
Function IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
IsInAp = HasEle(Av, V)
End Function

Function ExtzFfn$(Ffn)
ExtzFfn = Ext(Ffn)
End Function

Function Ext$(Ffn)
Dim B$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
Ext = Mid(B, P)
End Function

Function UpPth$(Pth, Optional NUp% = 1)
Dim O$: O = Pth
Dim J%
For J = 1 To NUp
    O = ParPth(O)
Next
UpPth = O
End Function

Function Pth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
Pth = Left(Ffn, P)
End Function

Function ParPthzFfn$(Ffn)
ParPthzFfn = ParPth(Pth(Ffn))
End Function

Function IsEqFfnStr(Ffn, S$) As Boolean
Dim L&: L = Len(S)
If FileLen(Ffn) <> L Then Exit Function
Dim J&, F%
F = FnoRnd128(Ffn)
For J = 1 To NBlk(SizFfn(Ffn), 128)
    Dim P&: P = (J - 1) * 128 + 1
    If FnoBlk(F, J) <> Mid(S, P, 128) Then
        Close #F
        Exit Function
    End If
Next
Close #F
IsEqFfnStr = True
End Function

Function IsEqFfn(A, B, Optional FilCmp As EmFilCmp = EmFilCmp.EiCmpEq) As Boolean
ThwIf_NoFfn A, CSub, "Fst File"
If A = B Then Thw CSub, "Fil A and B are eq name", "A", A
ThwIf_NoFfn B, CSub, "Snd File"
If Not IsSamFfn(A, B) Then Exit Function
If FilCmp = EiCmpSam Then
    IsEqFfn = True
    Exit Function
End If
Dim J&, F1%, F2%
F1 = FnoRnd128(A)
F2 = FnoRnd128(B)
For J = 1 To NBlk(SizFfn(A), 128)
    If FnoBlk(F1, J) <> FnoBlk(F2, J) Then
        Close #F1, F2
        Exit Function
    End If
Next
Close #F1, F2
IsEqFfn = True
End Function

Function IsSamFfn(A, B) As Boolean
If DtezFfn(A) <> DtezFfn(B) Then Exit Function
If Not IsSamzSi(A, B) Then Exit Function
IsSamFfn = True
End Function

Function IsSamzSi(Ffn1, Ffn2) As Boolean
IsSamzSi = SizFfn(Ffn1) = SizFfn(Ffn2)
End Function

Function MsgSamFfn(A, B, Si&, Tim$, Optional Msg$) As String()
Dim O$()
Push O, "File 1   : " & A
Push O, "File 2   : " & B
Push O, "File Size: " & Si
Push O, "File Time: " & Tim
Push O, "File 1 and 2 have same size and time"
If Msg <> "" Then Push O, Msg
MsgSamFfn = O
End Function

Sub Z_FfnBlk()
Dim T$, S$, A$
S = "sllksdfj lsdkjf skldfj skldfj lk;asjdf lksjdf lsdkfjsdkflj "
T = TmpFt
WrtStr S, T
Debug.Assert SizFfn(T) = Len(S)
A = FfnBlk(T, 1)
Debug.Assert A = Left(S, 128)
End Sub

Function FnoBlk$(Fno%, IBlk)
Dim A As String * 128
Get #Fno, IBlk, A
FnoBlk = A
End Function

Function FfnBlk$(Ffn, IBlk)
Dim F%: F = FnoRnd(Ffn, 128)
FfnBlk = FnoBlk(F, IBlk)
Close #F
End Function
Sub ThwIf_NotFxa(Ffn, Optional Fun$)
If Not IsFxa(Ffn) Then Thw Fun, "Given Ffn is not Fxa", "Ffn", Ffn
End Sub
Function IsFxa(Ffn) As Boolean
IsFxa = LCase(Ext(Ffn)) = FxaExt
End Function
Function IsFba(Ffn) As Boolean
IsFba = LCase(Ext(Ffn)) = FbaExt
End Function
Function IsPjf(Ffn) As Boolean
Select Case True
Case IsFxa(Ffn), IsFba(Ffn): IsPjf = True
End Select
End Function
Function IsFb(Ffn) As Boolean
Select Case LCase(Ext(Ffn))
Case FbExt, FbExt1: IsFb = True
End Select
End Function

Function IsFx(Ffn) As Boolean
Select Case LCase(Ext(Ffn))
Case ".xls", ".xlsm", ".xlsx": IsFx = True
End Select
End Function
Function FxAyzFfnAy(FfnAy$()) As String()
Dim Ffn
For Each Ffn In Itr(FfnAy)
    If IsFx(Ffn) Then PushI FxAyzFfnAy, Ffn
Next
End Function

Function FbAyzFfnAy(FfnAy$()) As String()
Dim Ffn
For Each Ffn In Itr(FfnAy)
    If IsFb(Ffn) Then PushI FbAyzFfnAy, Ffn
Next
End Function

Sub AsgExiMis(FfnAy$(), OExi$(), OMis$())
Dim Ffn
Erase OExi
Erase OMis
For Each Ffn In Itr(FfnAy)
    If HasFfn(Ffn) Then
        PushI OExi, Ffn
    Else
        PushI OMis, Ffn
    End If
Next
End Sub

Function HasFfn(Ffn) As Boolean
HasFfn = Fso.FileExists(Ffn)
End Function

Function NoFfn(Ffn) As Boolean
If Not HasFfn(Ffn) Then Debug.Print "NoFfn: " & Ffn: NoFfn = True
End Function

Function ExiFfnAset(FfnAy$()) As Aset
Set ExiFfnAset = AsetzAy(FfnAywExi(FfnAy))
End Function

Function MisFfnAset(FfnAy$()) As Aset
Set MisFfnAset = AsetzAy(FfnAywMis(FfnAy))
End Function

Function FfnAywExi(FfnAy$()) As String()
Dim F: For Each F In Itr(FfnAy)
    If HasFfn(F) Then PushI FfnAywExi, F
Next
End Function
Function FfnAywMis(FfnAy$()) As String()
Dim F: For Each F In Itr(FfnAy)
    If NoFfn(F) Then PushI FfnAywMis, F
Next
End Function

Sub ThwIf_NoEle(Ay, Ele, Fun$)
If NoEle(Ay, Ele) Then Thw Fun, "No @Ele in @Ay", "Ele Ay", Ele, Ay
End Sub

Sub ThwIf_NoFfn(Ffn, Fun$, Optional FilKind$)
If NoFfn(Ffn) Then Thw Fun, "File not found", "File-Pth File-Name File-Kind", Pth(Ffn), Fn(Ffn), FilKind
End Sub

Sub ThwIf_FfnExist(Ffn, Fun$, Optional FilKind$)
If HasFfn(Ffn) Then Thw Fun, "File already exist", "File-Pth File-Name File-Kind", Pth(Ffn), Fn(Ffn), FilKind
End Sub

Function RplExt$(Ffn, NewExt)
RplExt = RmvExt(Ffn) & NewExt
End Function

Function DtezFfn(Ffn) As Date
If HasFfn(Ffn) Then DtezFfn = FileDateTime(Ffn)
End Function

Function SizFfn&(Ffn)
If NoFfn(Ffn) Then SizFfn = -1: Exit Function
SizFfn = FileLen(Ffn)
End Function

Function SiDotDTim$(Ffn)
If HasFfn(Ffn) Then SiDotDTim = TimStr(DtezFfn(Ffn)) & "." & SizFfn(Ffn)
End Function

Function TimStrzFfn$(Ffn)
TimStrzFfn = TimStr(DtezFfn(Ffn))
End Function

Function AddTimSfxzFfn$(Ffn)
AddTimSfxzFfn = AddFnSfx(Ffn, Format(Now, "(HHMMSS)"))
End Function
Function AddFnPfx$(A$, Pfx$)
AddFnPfx = Pth(A) & Pfx & Fn(A)
End Function

Function AddFnSfx$(Ffn, Sfx$)
AddFnSfx = RmvExt(Ffn) & Sfx & Ext(Ffn)
End Function


Function NxtNozFfn%(Ffn)
Dim A$: A = Right(RmvExt(Ffn), 5)
If FstChr(A) <> "(" Then Exit Function
If LasChr(A) <> ")" Then Exit Function
Dim M$: M = Mid(A, 2, 3)
If Not IsDigStr(M) Then Exit Function
NxtNozFfn = M
End Function
Function RmvNxtNo$(Ffn)
If IsNxtFfn(Ffn) Then
    Dim A$: A = RmvExt(Ffn)
    RmvNxtNo = RmvLasNChr(A, 5) & Ext(Ffn)
Else
    RmvNxtNo = Ffn
End If
End Function
Sub Z_NxtFfn()
Dim Ffn$
'GoSub T0
GoSub T1
Exit Sub
T1: Ffn = "AA(000).xls"
    Ept = "AA(001).xls"
    GoTo Tst
T0:
    Ffn = "AA.xls"
    Ept = "AA(001).xls"
    GoTo Tst
Tst:
    Act = NxtFfn(Ffn)
    C
    Return
End Sub
Function NxtFfn$(Ffn)
Dim J&: J = NxtNozFfn(Ffn)
Dim F$: F = RmvNxtNo(Ffn)
NxtFfn = AddFnSfx(F, "(" & Pad0(J + 1, 3) & ")")
End Function
Function NxtFfnzNotIn(Ffn, NotInFfnAy$())
Dim J%, O$
O = Ffn
While HasEleS(NotInFfnAy, O)
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    O = NxtFfn(O)
Wend
NxtFfnzNotIn = O
End Function

Function NxtFfnzAva$(Ffn)
Dim J%, O$
O = Ffn
While HasFfn(O)
    If J = 999 Then Thw CSub, "Too much next file in the path of given-ffn", "Given-Ffn", Ffn
    J = J + 1
    O = NxtFfn(O)
Wend
NxtFfnzAva = O
End Function

Function NxtFfnAy(Ffn) As String() 'Return ffn and all it nxt ffn in the pth of given ffn
If HasFfn(Ffn) Then Push NxtFfnAy, Ffn  '<==
Dim A$()
    Dim Spec$
        Spec = AddFnSfx(Fn(Ffn), "(???)")
    A = FfnAy(Pth(Ffn), Spec)
Dim I, F$
For Each I In Itr(A)
    F = I
    If IsNxtFfn(Ffn) Then PushI NxtFfnAy, F   '<==
Next
End Function

Function IsNxtFfn(Ffn) As Boolean
Select Case True
Case NxtNozFfn(Ffn) > 0, Right(Fnn(Ffn), 5) = "(000)": IsNxtFfn = True
End Select
End Function


Function InstFfn$(Ffn)
InstFfn = InstPth(Pth(Ffn)) & Fn(Ffn)
End Function

Function InstPth$(Pth)
InstPth = AddFdrEns(Pth, NowStr)
End Function

Function InstFdr$(Fdr)
InstFdr = AddFdrEns(TmpFdr(Fdr), NowStr)
End Function
Function CrtPthzInst$(Pth)
CrtPthzInst = InstPth(Pth)
End Function

Function IsInstFfn(Ffn) As Boolean
IsInstFfn = IsInstFdr(FdrzFfn(Ffn))
End Function

Function IsInstFdr(Fdr$) As Boolean
IsInstFdr = IsTimStr(Fdr)
End Function

Function FfnzPthFn$(Pth, Fn$)
FfnzPthFn = Ffn(Pth, Fn)
End Function

Function Ffn$(Pth, Fn$)
Ffn = EnsPthSfx(Pth) & Fn
End Function

Function HasExtss(Ffn, ExtssLin) As Boolean
Dim E$: E = Ext(Ffn)
Dim Sy$(): Sy = SyzSS(ExtssLin)
HasExtss = HasEleS(Sy, E)
End Function


