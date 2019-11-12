Attribute VB_Name = "MxDoMthLis"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMthLis."
Enum EmTriSte
    EiTriOpn
    EiTriYes
    EiTriNo
End Enum
Public Const FFoMthLis$ = "Pjn MdTy Mdn L Mdy Ty Mthn TyChr RetAs ShtPm"
':JSrc: :Lin #Jmp-Lin# ! Fmt: T1 Rst, *T1 is JmpLin"<mdn:Lno>".  *Rst is '<SrcLin>

Sub LisPj()
Dim A$()
    A = PjNyzV(CVbe)
    D AmAddPfx(A, "ShwPj """)
D A
End Sub

Sub LisStopLin()
LisSrc "Stop"
End Sub

Sub LisPubPatn(Patn$)
Dim A As Drs: A = DoFun
BrwDrs DwPatn(A, "Mthn", Patn)
End Sub


Sub LisPFunRetAs(RetAsPatn$)
Dim RetSfx As Drs: RetSfx = AddMthColRetAs(DoPubFun)
Dim Patn As Drs: Patn = DwPatn(RetSfx, "RetSfx", RetAsPatn)
Dim T50 As Drs: T50 = DwTopN(Patn)
BrwDrs T50
End Sub

Sub LisPPrpRetAs(RetAsPatn$)
Dim S As Drs: S = DoPubFun
Dim RetSfx As Drs: RetSfx = AddMthColRetAs(S)
Dim Pub As Drs: Pub = DwEqExl(RetSfx, "Mdy", "Pub")
Dim Fun As Drs: Fun = DwEqExl(Pub, "Ty", "Get")
Dim Patn As Drs: Patn = DwPatn(Fun, "RetSfx", RetAsPatn)
Dim T50 As Drs: T50 = DwTopN(Patn)
BrwDrs T50
End Sub

Sub LisMthCntzQIde()
DmpDrs SrtDrs(DwEq(DoMthCntP, "Lib", "QIde")), Fmt:=EiSSFmt, IsSum:=True
End Sub

Function JSrc$(Mdn$, Lno&, C1%, C2%, Lin)
JSrc = FmtQQ("JmpLin""?:?:?:?"" '?", Mdn, Lno, C1, C2, Lin)
End Function

Function JSrczPred(P As IPred) As String()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    Dim Md As CodeModule: Set Md = C.CodeModule
    Dim L, Lno&: Lno = 0
    Dim C1%, C2%
    If Md.CountOfLines > 0 Then
        For Each L In Itr(SplitCrLf(Md.Lines(1, Md.CountOfLines)))
            Lno = Lno + 1
            If P.Pred(L) Then
                C1 = 1
                C2 = 1
                PushI O, JSrc(C.Name, Lno, C1, C2, L)
            End If
        Next
    End If
Next
JSrczPred = AlignLyz1T(O)
End Function

Function JSrczIdf(Idf$) As String()
JSrczIdf = JSrczPred(PredHasIdf(Idf))
End Function

Function JSrczPfx(LinPfx$) As String()
JSrczPfx = JSrczPred(PredHasPfx(LinPfx))
End Function

Function JSrczPatn(LinPatn$, Optional AndPatn1$, Optional AndPatn2$) As String()
JSrczPatn = JSrczPred(PredHasPatn(LinPatn, AndPatn1, AndPatn2))
End Function

Sub LisSrcoPfx(LinPfx$, Optional OupTy As EmOupTy = EmOupTy.EiOtDmp)
Brw JSrczPfx(LinPfx), OupTy:=OupTy
End Sub

Sub LisSrcoIdf(Idf$, Optional OupTy As EmOupTy = EmOupTy.EiOtDmp)
':Idf: :Nm #Identifier#
Brw JSrczIdf(Idf), OupTy:=OupTy
End Sub

Sub LisSrc(LinPatn$, Optional AndPatn1$, Optional AndPatn2$, Optional OupTy As EmOupTy = EmOupTy.EiOtDmp)
Brw JSrczPatn(LinPatn, AndPatn1, AndPatn2), OupTy:=OupTy
End Sub
Sub AsgPatn123(PatnSS3$, OPatn$, OPatn1$, OPatn2$)
Dim A$(): A = SyzSS(PatnSS3)
If Si(A) > 3 Then Thw CSub, "PatnSS3 should have 3 or less SSTerm", "PatnSS3", PatnSS3
AsgSS PatnSS3, OPatn, OPatn1, OPatn2
End Sub

Sub AsgSS(SS$, ParamArray OAp())
Dim A$(): A = SyzSS(SS)
Dim Av(): Av = OAp
Dim U1%, U2%: U1 = UB(A): U2 = UB(Av)
Dim J%
For J = 0 To U2: OAp(J) = Empty: Next
For J = 0 To Min(U1, U2)
    OAp(J) = A(J)
Next
End Sub

Sub LisMthzP(Optional PatnSS3$, Optional ExlPatn$, Optional ShtMdySS$, Optional ShtMthTySS$, _
Optional TyChr$, Optional RetAsPatn$, Optional ShouldRetAy As EmTriSte, _
Optional NPm% = -1, Optional ShtPmPatn$, Optional HasAp As EmTriSte, _
Optional MdnPatn$, Optional ShtMdTySS$, _
Optional OupTy As EmOupTy, Optional Top% = 50)
Dim P$, P1$, P2$: AsgPatn123 PatnSS3, P, P1, P2
LisMth P, P1, P2, ShtMdySS, ExlPatn, ShtMthTySS, _
        TyChr, RetAsPatn, ShouldRetAy, _
        NPm, ShtPmPatn, HasAp, _
        MdnPatn, ShtMdTySS
End Sub

Sub LisMth(Optional Patn$, Optional Patn1$, Optional Patn2$, Optional ExlPatn$, Optional ShtMdySS$, Optional ShtMthTySS$, _
Optional TyChr$, Optional RetAsPatn$, Optional ShouldRetAy As EmTriSte, _
Optional NPm% = -1, Optional ShtPmPatn$, Optional HasAp As EmTriSte, _
Optional MdnPatn$, Optional ShtMdTySS$, _
Optional OupTy As EmOupTy, Optional Top% = 50)
Dim D As Drs:
    D = DwDoMthLis(DoMthLisP, Patn, Patn1, Patn2, ShtMdySS, ShtMthTySS, _
        TyChr, RetAsPatn, ShouldRetAy, _
        NPm, ShtPmPatn, HasAp, _
        MdnPatn, ShtMdTySS)
    D = DePatn(D, "Mthn", ExlPatn)
Dim D1 As Drs: D1 = DwTopN(D, Top)
Brw FmtDrs(D1, , , , EiBeg1, EiSSFmt), OupTy:=OupTy
End Sub

Function PatnzSS$(SS, LisAy$())
Dim A$(): A = AwDist(SyzSS(SS))
Dim b$()
    Dim I: For Each I In Itr(A)
        If HasEle(LisAy, I) Then
            PushNDup b, I
        End If
    Next
Dim C$: C = Jn(b, "|")
If C = "" Then Exit Function
PatnzSS = Qte(C, "()")
End Function

Function AAAA()
'ß
End Function

Function DoMthLisP() As Drs
DoMthLisP = DoMthLiszP(CPj)
End Function

Function DoMthLiszP(P As VBProject) As Drs
DoMthLiszP = DoMthLis(DoMthzP(P))
End Function

Function DoMthLis(DoMth As Drs) As Drs
DoMthLis = SelDrs(Add5MthCol(DoMth), FFoMthLis)
End Function

Function DwDoMthLis(DoMthLis As Drs, Patn$, Patn1$, Patn2$, ShtMdySS$, ShtMthTySS$, _
TyChr$, RetAsPatn$, RetAy As EmTriSte, _
NPm%, ShtPmPatn$, HasAp As EmTriSte, _
MdnPatn$, ShtMdTySS$) As Drs

'- Pfx-Pn = Patn
Dim PnMdy$:             PnMdy = PatnzSS(ShtMdySS, ShtMthMdyAy)
Dim PnTy$:               PnTy = PatnzSS(ShtMthTySS, ShtMthTyAy)
'- Pfx-I = Inp-Do-Fm-DoMthLis
Dim IMdy     As Drs:     IMdy = DwPatn(DoMthLisP, "Mdy", PnMdy)
Dim ITy      As Drs:      ITy = DwPatn(IMdy, "Ty", PnTy)
Dim ITyChr As Drs:     ITyChr = DwEqStr(ITy, "TyChr", TyChr)
Dim IPatn    As Drs:    IPatn = DwPatn(ITyChr, "Mthn", Patn, Patn1, Patn2)
Dim IHasAp   As Drs:   IHasAp = DwHasAp(IPatn, HasAp)
Dim INPm     As Drs:     INPm = DwNPm(IHasAp, NPm)
Dim IMdn     As Drs:     IMdn = DwPatn(INPm, "Mdn", MdnPatn)
Dim IRetAs   As Drs:   IRetAs = DwPatn(IMdn, "RetAs", RetAsPatn)
Dim IRetAy   As Drs:   IRetAy = DwRetAy(IRetAs, RetAy)
                   DwDoMthLis = DwPatn(IRetAy, "ShtPm", ShtPmPatn)
End Function

Function DwRetAy(WiRetAs As Drs, RetAy As EmTriSte) As Drs
If RetAy = EiTriOpn Then DwRetAy = WiRetAs: Exit Function
Dim RetAy1 As Boolean: RetAy1 = BoolzTriSte(RetAy)
Dim IRetAs%: IRetAs = IxzAy(WiRetAs.Fny, "RetAs")
Dim ODy()
    Dim Dr: For Each Dr In Itr(WiRetAs.Dy)
        Dim RetAs$: RetAs = Dr(IRetAs)
        If HasSfx(RetAs, "()") = RetAy1 Then PushI ODy, Dr
    Next
DwRetAy = Drs(WiRetAs.Fny, ODy)
End Function

Function HasAp(MthPm) As Boolean
Dim A$(): A = SplitCommaSpc(MthPm): If Si(A) = 0 Then Exit Function
HasAp = HasPfx(LasEle(A), "ParamArray ")
End Function

Function BoolzTriSte(A As EmTriSte) As Boolean
Select Case True
Case A = EiTriYes: BoolzTriSte = True
Case A = EiTriNo:  BoolzTriSte = False
Case Else: Stop
End Select
End Function

Function DwHasAp(WiMthPm As Drs, HasAp0 As EmTriSte) As Drs
If HasAp0 = EiTriOpn Then DwHasAp = WiMthPm: Exit Function
Dim HasAp1 As Boolean: HasAp1 = BoolzTriSte(HasAp0)
Dim IMthPm%: IMthPm = IxzAy(WiMthPm.Fny, "MthPm")
Dim ODy()
    Dim Dr: For Each Dr In Itr(WiMthPm.Dy)
        Dim MthPm$: MthPm = Dr(IMthPm)
        If HasAp1 = HasAp(MthPm) Then PushI ODy, Dr
    Next
DwHasAp = Drs(WiMthPm.Fny, ODy)
End Function

Function DwNPm(D As Drs, NPm%) As Drs
If NPm < 0 Then DwNPm = D: Exit Function
Dim Ix%: Ix = IxzAy(D.Fny, "MthPm")
Dim ODy(), Dr, Pm$: For Each Dr In Itr(D.Dy)
    Pm = Dr(Ix)
    If Si(SplitComma(Pm)) = NPm Then PushI ODy, Dr
Next
DwNPm = Drs(D.Fny, ODy)
End Function

