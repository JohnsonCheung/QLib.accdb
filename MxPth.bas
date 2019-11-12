Attribute VB_Name = "MxPth"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Pth"
Const CMod$ = CLib & "MxPth."

':Pseg: :S #Pth-Segment# ! 0 or more :Fdr separated by :PthSep
':Fdr:  :S #Folder#      ! An directory entry in file-system
Function IsPseg(Pseg$) As Boolean
Select Case True
Case FstChr(Pseg) = "\"
Case LasChr(Pseg) = "\"
Case Else: IsPseg = True
End Select
End Function

Function AddFdr$(Pth, Fdr)
AddFdr = EnsPthSfx(Pth) & AddNB(Fdr, "\")
End Function

Function AddPseg$(Pth, Pseg)
'Ret : :Pth
Dim O$: O = Pth
    Dim Fdr: For Each Fdr In Itr(Split(Pseg, PthSep))
        If Fdr <> "" Then O = O & Fdr & PthSep
    Next
AddPseg = EnsPthSfx(O)
End Function

Function AddPsegEns$(Pth, Pseg)
AddPsegEns = EnsPthAll(AddPseg(Pth, Pseg))
End Function

Function AddFdrEns$(Pth, Fdr)
AddFdrEns = EnsPth(AddFdr(Pth, Fdr))
End Function

Function AddFdrApEns$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
Dim O$: O = AddFdrAv(Pth, Av)
EnsPthAll O
AddFdrApEns = O
End Function

Function AddFdrAv$(Pth, FdrAv())
Dim O$: O = Pth
Dim I, Fdr$
For Each I In FdrAv
    Fdr = I
    O = AddFdr(O, Fdr)
Next
AddFdrAv = O
End Function

Function AddFdrAp$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
AddFdrAp = AddFdrAv(Pth, Av)
End Function

Function IsEmpPth(Pth) As Boolean
ThwIf_NoPth Pth, CSub
If AnyFil(Pth) Then Exit Function
If HasSubFdr(Pth) Then Exit Function
IsEmpPth = True
End Function

Function AddPfxzPth$(Pth, Pfx)
With Brk2Rev(RmvSfx(Pth, PthSep), PthSep, NoTrim:=True)
    AddPfxzPth = .S1 & PthSep & Pfx & .S2 & PthSep
End With
End Function

Function HitFilAtr(A As VbFileAttribute, Wh As VbFileAttribute) As Boolean
HitFilAtr = True
End Function

Function FdrzFfn$(Ffn)
FdrzFfn = Fdr(Pth(Ffn))
End Function

Function Fdr$(Pth)
Fdr = AftRev(DltPthSfx(Pth), PthSep)
End Function

Sub ThwIf_NotProperFdrNm(Fdr$)
Const CSub$ = CMod & "ThwNotFdr"
Const C$ = "\/:<>"
If HasChrList(Fdr, C) Then Thw CSub, "Fdr cannot has these char " & C, "Fdr Char", Fdr, C
End Sub

Function RmvFdr$(Pth)
RmvFdr = BefRev(DltPthSfx(Pth), PthSep) & PthSep
End Function

Function ParPth$(Pth) ' Return the ParPth of given Pth
ParPth = RmvFdr(Pth)
End Function

Function ParFdr$(Pth)
ParFdr = Fdr(ParPth(Pth))
End Function

Function PthzUpN$(Pth, UpN%)
Dim O$: O = Pth
Dim J%: For J = 1 To UpN
    O = ParPth(O)
Next
PthzUpN = O
End Function

Function EnsPth$(Pth)
Dim P$: P = EnsPthSfx(Pth)
If NoPth(P) Then MkDir RmvLasChr(P)
EnsPth = P
End Function

Function EnsPthAll$(Pth)
'Ret :Pth and ens each :Pseg. @@
Dim J%, O$, Ay$()
Ay = Split(RmvSfx(Pth, PthSep), PthSep)
O = Ay(0)
For J = 1 To UBound(Ay)
    O = O & PthSep & Ay(J)
    EnsPth O
Next
EnsPthAll = Pth
End Function

Function HasPth(Pth) As Boolean
HasPth = Fso.FolderExists(Pth)
End Function

Function NoPth(Pth) As Boolean
If Not HasPth(Pth) Then Debug.Print "NoPth: "; Pth: NoPth = True
End Function

Function HasFdr(Pth, Fdr$) As Boolean
HasFdr = HasEle(FdrAy(Pth), Fdr)
End Function

Sub ThwIf_NoPth(Pth, Fun$)
If NoPth(Pth) Then Thw Fun, "Pth not exist", "Pth", Pth
End Sub

Function AnyFil(Pth) As Boolean
AnyFil = Dir(Pth) <> ""
End Function

Function HasSubFdr(Pth) As Boolean
HasSubFdr = Fso.GetFolder(Pth).SubFolders.Count > 0
End Function

Function DirzPSA$(Pth, Optional Spec$ = "*.*", Optional A As VbFileAttribute = VbFileAttribute.vbDirectory)
DirzPSA = Dir(EnsPthSfx(Pth) & Spec, A)
End Function

Function FdrAyzIsInst(Pth) As String()
Dim I, Fdr$
For Each I In Itr(FdrAy(Pth))
    Fdr = I
    If IsInstNm(Fdr) Then PushI FdrAyzIsInst, Fdr
Next
End Function

Function FdrAy(Pth, Optional Spec$ = "*.*") As String()
If NoPth(Pth) Then Exit Function
Dim P$: P = EnsPthSfx(Pth)
Dim F, X&, Atr1&
For Each F In Itr(EntAy(P, Spec))
    Atr1 = GetAttr(P & F)
    If (Atr1 And VbFileAttribute.vbDirectory) <> 0 Then
        PushI FdrAy, F    '<====
    End If
Next
End Function

Function EntAy(Pth, Optional Spec$ = "*.*", Optional AtR As FileAttribute = vbDirectory) As String()
Dim A$: A$ = DirzPSA(EnsPthSfx(Pth), Spec, AtR)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If InStr(A, "?") > 0 Then
        Inf CSub, "Unicode entry is skipped", "UniCode-Entry Pth Spec", A, Pth, Spec
        GoTo X
    End If
    PushI EntAy, A
X:
    A = Dir
Wend
End Function
Function IsInstNm(Nm) As Boolean
If FstChr(Nm) <> "N" Then Exit Function      'FstChr = N
If Len(Nm) <> 16 Then Exit Function          'Len    =16
If Not IsYYYYMMDD(Mid(Nm, 2, 8)) Then Exit Function 'NYYYYMMDD_HHMMDD
If Mid(Nm, 10, 1) <> "_" Then Exit Function
If Not IsHHMMDD(Right(Nm, 6)) Then Exit Function
IsInstNm = True
End Function

Function FdrAy1(Pth) As String()
Dim P$: P = EnsPthSfx(Pth)
Dim A$: A = DirzPSA(P)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If HasPth(P & A) Then PushI FdrAy1, A
X:
    A = Dir
Wend
End Function

Function FfnItr(Pth)
Asg Itr(FfnAy(Pth)), FfnItr
End Function

Function SubPthAy(Pth) As String()
SubPthAy = AmAddPfxS(FdrAy(Pth), EnsPthSfx(Pth), PthSep)
End Function

Sub AsgEnt(OFdrAy$(), OFnAy$(), Pth)
Erase OFdrAy
Erase OFnAy
Dim A$, P$
P = EnsPthSfx(Pth)
A = Dir(Pth, vbDirectory)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If HasPth(P & A) Then
        PushI OFdrAy, A
    Else
        PushI OFnAy, A
    End If
    A = Dir
X:
Wend
End Sub

Function FnnAy(Pth, Optional Spec$ = "*.*") As String()
Dim I: For Each I In FnAy(Pth, Spec)
    PushI FnnAy, RmvExt(I)
Next
End Function

Function FnAyzFfnAy(FfnAy$()) As String()
Dim I, Ffn$
For Each I In Itr(FfnAy)
    Ffn = I
    PushI FnAyzFfnAy, Fn(Ffn)
Next
End Function

Function FnAy(Pth, Optional Spec$ = "*.*") As String()
Dim O$()
Dim M$: M = Dir(EnsPthSfx(Pth) & Spec)
While M <> ""
   PushI FnAy, M
   M = Dir
Wend
End Function

Function Fxy(Pth) As String()
Fxy = FfnAy(Pth, "*.xls*")
End Function

Function FfnAy(Pth, Optional Spec$ = "*.*") As String()
FfnAy = AmAddPfx(FnAy(Pth, Spec), EnsPthSfx(Pth))
End Function

Sub Z_SubPthAy()
Dim Pth
Pth = "C:\Users\user\AppData\Local\Temp\"
Ept = Sy()
GoSub Tst
Exit Sub
Tst:
    Act = SubPthAy(Pth)
    Brw Act
    Return
End Sub

Sub Z_Fxy()
Dim A$()
A = Fxy(CurDir)
DmpAy A
End Sub

Function HasPthSfx(Pth) As Boolean
HasPthSfx = LasChr(Pth) = PthSep
End Function

Function EnsPthSfx$(Pth)
If Pth = "" Then Exit Function
If HasPthSfx(Pth) Then
    EnsPthSfx = Pth
Else
    EnsPthSfx = Pth & PthSep
End If
End Function

Function DltPthSfx$(Pth)
DltPthSfx = RmvSfx(Pth, PthSep)
End Function

Function HasSiblingFdr(Pth, Fdr$) As Boolean
HasSiblingFdr = HasFdr(ParPth(Pth), Fdr)
End Function

Function SiblingPth$(Pth, SiblingFdr$)
SiblingPth = AddFdrEns(ParPth(Pth), SiblingFdr)
End Function
