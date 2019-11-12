Attribute VB_Name = "MxDoMthCnt"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMthCnt."
Const CntgMthPP$ = "NPubSub NPubFun NPubPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp"
Enum EmOupTy
    EiOtDmp
    EiOtBrw
    EiOtVc
End Enum
Type CntgMth
    Lib As String
    Mdn As String
    NPubSub As Integer
    NPubFun As Integer
    NPubPrp As Integer
    NPrvSub As Integer
    NPrvFun As Integer
    NPrvPrp As Integer
    NFrdSub As Integer
    NFrdFun As Integer
    NFrdPrp As Integer
End Type
Type CntgMths: N As Long: Ay() As CntgMth: End Type

Function DoMthCntP(Optional MdnPatn$ = ".+", Optional SrtCol$ = "Mdn") As Drs
DoMthCntP = DoMthCntzP(CPj, MdnPatn, SrtCol)
End Function

Function DoMthCntzP(P As VBProject, MdnPatn$, SrtCol$) As Drs
Dim R As RegExp: Set R = Rx(MdnPatn, IgnoreCase:=True)
Dim C As VBComponent, Dy(): For Each C In P.VBComponents
    If R.Test(C.Name) Then
        PushI Dy, DroMthCnt(C.CodeModule)
    End If
Next
Dim D As Drs: D = Drs(FoMthCnt, Dy)
DoMthCntzP = SrtDrs(D, SrtCol)
End Function
Function FoMthCnt() As String()
FoMthCnt = SyzSS("Lib Mdn NLines NMth NPubSub NPubFun NPubPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp")
End Function
Function NMth%(A As CntgMth)
With A
NMth = .NPubSub + .NPubFun + .NPubPrp + .NPrvSub + .NPrvFun + .NPrvPrp + .NFrdSub + .NFrdFun + .NFrdPrp
End With
End Function
Function FmtCntgMth(A As CntgMth, Optional Hdr As EmHdr)
Dim Pfx$: If Hdr = EiWiHdr Then Pfx = "Pub* | Prv* | Frd* : *{Sub Fun Frd} "
With A
Dim N%: N = NMth(A)
FmtCntgMth = JnAp(" | ", Pfx, .Mdn, N) & " | " & JnSpcAp(.NPubSub, .NPubFun, .NPubPrp, .NPrvSub, .NPrvFun, .NPrvPrp, .NFrdSub, .NFrdFun, .NFrdPrp)
End With
End Function

Function NMthzS%(Src$())
NMthzS = Si(MthIxy(Src))
End Function

Function NMthP%()
NMthP = NMthzP(CPj)
End Function

Function NMthM%()
NMthM = NMthzM(CMd)
End Function

Function NMthzP%(Pj As VBProject)
Dim O%, C As VBComponent
For Each C In Pj.VBComponents
    O = O + NMthzS(Src(C.CodeModule))
Next
NMthzP = O
End Function

Function DroMthCnt(M As CodeModule) As Variant()
Dim S$(): S = Src(M)
Dim Mth$(): Mth = MthLinAy(S)
Dim L: For Each L In Itr(Mth)
    With Mthn3zL(L)
        Dim Prv As Boolean: Prv = False
        Dim Pub As Boolean: Pub = False
        Dim Frd As Boolean: Frd = False
        Dim Fun As Boolean: Fun = False
        Dim Sbr As Boolean: Sbr = False
        Dim Prp As Boolean: Prp = False
        Select Case .ShtMdy
        Case "Prv": Prv = True
        Case "Pub": Pub = True
        Case "Frd": Frd = True
        Case Else: Thw CSub, "Out of valid value: Prv PUb Frd", "ShtMdy", .ShtMdy
        End Select
        Select Case ShtMthKdzShtMthTy(.ShtTy)
        Case "Fun": Fun = True
        Case "Sub": Sbr = True
        Case "Prp": Prp = True
        Case Else: Thw CSub, "Out of valid value: Sub Fun Prp", "ShtMdy", .ShtMdy
        End Select
    End With
            
    Select Case True
        Case Pub And Sbr: Dim NPubSub%: NPubSub = NPubSub + 1
        Case Pub And Fun: Dim NPubFun%: NPubFun = NPubFun + 1
        Case Pub And Prp: Dim NPubPrp%: NPubPrp = NPubPrp + 1
        Case Prv And Sbr: Dim NPrvSub%: NPrvSub = NPrvSub + 1
        Case Prv And Fun: Dim NPrvFun%: NPrvFun = NPrvFun + 1
        Case Prv And Prp: Dim NPrvPrp%: NPrvPrp = NPrvPrp + 1
        Case Frd And Sbr: Dim NFrdSub%: NFrdSub = NFrdSub + 1
        Case Frd And Fun: Dim NFrdFun%: NFrdFun = NFrdFun + 1
        Case Frd And Prp: Dim NFrdPrp%: NFrdPrp = NFrdPrp + 1
        Case Else: Thw CSub, "Invalid Mthn3", "MthLin", L
    End Select
    Dim NMth%: NMth = NMth + 1
    If NPubSub + NPubFun + NPubPrp + NPrvSub + NPrvFun + NPrvPrp + NFrdSub + NFrdFun + NFrdPrp <> NMth Then Stop
Next
Dim NLin&: NLin = Si(S)
Dim Mdn$: Mdn = MdnzM(M)
Dim Lib$: Lib = Bef(Mdn, "_")
DroMthCnt = Array(Lib, Mdn, NLin, NMth, NPubSub, NPubFun, NPubPrp, NPrvSub, NPrvFun, NPrvPrp, NFrdSub, NFrdFun, NFrdPrp)
End Function

Sub LisMdP(Optional MdnPatn$ = ".+", Optional SrtCol$ = "Mdn", Optional OupTy As EmOupTy)
LisMdzP CPj, MdnPatn, SrtCol, OupTy
End Sub

Sub LisMdzM(M As CodeModule)
DmpDrs DoMthCntzM(M)
End Sub

Function DoMthCntzM(M As CodeModule) As Drs
DoMthCntzM = Drs(FoMthCnt, Av(DroMthCnt(M)))
End Function

Sub CntWrdP()
Debug.Print WrdStszLy(SrczP(CPj))
End Sub

Sub LisMdM()
LisMdzM CMd
End Sub

Sub LisMdzP(P As VBProject, MdnPatn$, SrtCol$, OupTy As EmOupTy)
Brw FmtDrszNoRdu(DoMthCntzP(P, MdnPatn, SrtCol), Fmt:=EiSSFmt), OupTy:=OupTy
End Sub

Function NMthzM%(M As CodeModule)
NMthzM = NMthzS(Src(M))
End Function

Function NSrcLinPj&(P As VBProject)
Dim O&, C As VBComponent
For Each C In P.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinPj = O
End Function

Function NPubMthzS%(Src$())
NPubMthzS = NItr(PubMthLinItr(Src))
End Function

Function NPubMthzM%(M As CodeModule)
NPubMthzM = NPubMthzS(Src(M))
End Function

Function NPubMthzV%(A As Vbe)
Dim O%, P As VBProject
For Each P In A.VBProjects
    O = O + NPubMthzP(P)
Next
NPubMthzV = O
End Function

Property Get NPubMthV%()
NPubMthV = NPubMthzV(CVbe)
End Property

Function NPubMthzP%(P As VBProject)
Dim O%, C As VBComponent
For Each C In P.VBComponents
    O = O + NPubMthzM(C.CodeModule)
Next
NPubMthzP = O
End Function
