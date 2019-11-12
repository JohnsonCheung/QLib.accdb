Attribute VB_Name = "MxMthn"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthn."
':Mthn: :Nm ! Rule1-FstVerbBeingDo: the mthn will not return any value
'       ! Rule2-FstVerbBeingDo: tThe Cmls aft Do is a verb
':Dta_MthQVNm: :Nm ! It is a String dervied from Nm.  Q for quoted.  V for verb.  It has 3 Patn: NoVerb-[#xxx], MidVerb-[xxx(vvv)xxx], FstVerb-[(vvv)xxx]."
':Nm: :S ! less that 64 chr.
':FunNm: :Rul ! If there is a Subj in pm, put the Subj as fst CmlTerm and return that Subj;
'       ! give a Noun to the subj noun is MulCml.
'       ! Each Mthn must belong to one of these rule:
'       !   Noun | Noun.Verb.Extra | Verb.Variant | Noun.z.Variant
'       ! Pm-Rule
'       !   Subj    : Choose a subj in pm if there is more than one arg"
'       !   MuliNoun: It is Ok to group mul-arg as one subj
'       !   MulNounUseOneCml: Mul-noun as one subj use one Cml
':Noun: :Nm  ! it is 1 or more Cml to form a Noun."
':Cml:  :Nm  ! Tag:Type. P1.NumIsLCase:.  P2.LowDashIsLCase:.  P3.FstChrCanAnyNmChr:.
':Sfxss: :SS !  NmRul means variable or function name.
':VdtVerss: :SS ! P1.Opt: Each module may one DoczVdtSSoVerb.  P2.OneOccurance: "
':NounVerbExtra :SS !Tag: FunNmRule.  Prp1.TakAndRetNoun: Fst Cml is Noun and Return Noun.  Prp2.OneCmlNoun: Noun should be 1 Cml.  " & _
'                ! Prp3.VdtVerb: Snd Cml should be approved/valid noun.  Prp4.OptExtra: Extra is optional."
Sub AsgDNm(DNm$, O1$, O2$, O3$)
Dim Ay$(): Ay = Split(DNm, ".")
Select Case Si(Ay)
Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
Case Else: Stop
End Select
End Sub

Function MthDnzM$(M As CodeModule, Lin)
Dim D$: D = MthDnzL(Lin): If D = "" Then Exit Function
MthDnzM = MdDn(M) & "." & D
End Function

Function PubMthNy(Src$()) As String()
Dim Ix, N$, B As Mthn3
For Each Ix In MthIxItr(Src)
    B = Mthn3zL(Src(Ix))
    If B.Nm <> "" Then
        If B.ShtMdy = "" Or B.ShtMdy = "Pub" Then
            PushI PubMthNy, B.Nm
        End If
    End If
Next
End Function

Function Mthnn$()
Mthnn = MthnnzM(CMd)
End Function
Function MthnnzM$(M As CodeModule)
MthnnzM = MthnnzS(Src(M))
End Function

Function MthnnzS$(Src$())
MthnnzS = MthnnzL(MthLinAy(Src))
End Function
Function MthnnzL$(MthLinAy$())
MthnnzL = JnSpc(AySrt(MthNy(MthLinAy)))
End Function

Function MthNyzL(MthLinAy$()) As String()
Const CSub$ = CMod & "MthNyzL"
Dim I, Nm$, J%, MthLin
For Each I In Itr(MthLinAy)
    Nm = MthnzLin(I)
    If Nm = "" Then Thw CSub, "Given MthLinAy does not have Mthn", "[MthLin with error] Ix MthLinAy", I, J, AddIxPfx(MthLinAy)
    PushI MthNyzL, Nm
    J = J + 1
Next
End Function
Function Ens2Dot(S) As StrOpt
Select Case DotCnt(S)
Case 0: Ens2Dot = SomStr(".." & S)
Case 1: Ens2Dot = SomStr("." & S)
Case 2: Ens2Dot = SomStr(S)
End Select
End Function

Function RmvMthMdy$(Lin)
RmvMthMdy = RmvT1XzA(Lin, MthMdyAy)
End Function

Function Dimn$(Lin)
Dim L$: L = Lin
If ShfTerm(L, "Dim") Then Dimn = Nm(LTrim(L))
End Function

Function DimNy(Ly$()) As String()
Dim L
For Each L In Itr(Ly)
    PushI DimNy, Dimn(L)
Next
End Function

Function Mthn$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfMthTy(L) = "" Then Exit Function
Mthn = Nm(L)
End Function

Sub Z_MthDnzL()
Debug.Print MthDnzL("Function MthnzMthDn$(MthDn$)")
End Sub

Function MthKn$(Nm$, ShtTy$)
MthKn = Nm & "." & ShtTy
End Function

Function MthDn$(Nm$, ShtMdy$, ShtTy$)
MthDn = JnDotAp(Nm, ShtMdy, ShtTy)
End Function

Function MthDnzMthn3$(A As Mthn3)
MthDnzMthn3 = MthDn(A.Nm, A.ShtMdy, A.ShtTy)
End Function

Function MthDnzL$(Lin)
MthDnzL = MthDnzMthn3(Mthn3zL(Lin))
End Function

Function MthnzLin(Lin)
MthnzLin = Mthn(Lin)
End Function

Function PrpNm$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfKd(L) <> "Property" Then Exit Function
PrpNm = Nm(L)
End Function

Function MthnzDNm$(MthDn)
Dim Ay$(): Ay = Split(MthDn, ".")
Dim Nm$
Select Case Si(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
MthnzDNm = Nm
End Function
Sub Z_Mthn()
GoTo Z
Dim A$
A = "Function Mthn(A)": Ept = "Mthn.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = Mthn(A)
    C
    Return
Z:
    Dim O$(), L
    For Each L In SrczV(CVbe)
        PushNB O, Mthn(CStr(L))
    Next
    Brw O
End Sub

Function MthMdy$(Lin)
MthMdy = IFin(T1(Lin), MthMdyAy)
End Function

Function MthKd$(Lin)
MthKd = TakMthKd(RmvMdy(Lin))
End Function

Function Rpl$(S, SubStr$, By$, Optional Ith% = 1)
Dim P&: P = InStrWiIthSubStr(S, SubStr, Ith)
If P = 0 Then Rpl = S: Exit Function
Rpl = Replace(S, SubStr, By, P, 1)
End Function

Property Get Rel0Mthn2Mdn() As Rel
Dim O As New Rel
End Property

Function ModNyzPubMth(PubMthn) As String()
ModNyzPubMth = ModNyzPjPubMth(CPj, PubMthn)
End Function
Function PubMthNyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsLinPubMth(L) Then PushI PubMthNyzS, Mthn(L)
Next
End Function

Sub Z_ModNyzPjPubMth()
Dim P As VBProject, PubMthn
GoSub Z
Exit Sub
Z:
    D ModNyzPjPubMth(CPj, "AA")
    Stop
    Return
End Sub

Function HasPubMth(Src$(), PubMthn) As Boolean
Dim L: For Each L In Itr(Src)
    With Mthn3zL(L)
        If .ShtMdy = "Pub" Then
            If .Nm = PubMthn Then
                HasPubMth = True
                Exit Function
            End If
        End If
    End With
Next
End Function

Function ModNyzPjPubMth(P As VBProject, PubMthn) As String()
#If False Then
Dim Src As Drs: Src = DoPubFunzP(P)
Dim Sel As Drs: Sel = Dw2Eq(Src, "Mdy MdTy", "Pub", "Std")
#Else
Dim I, M As CodeModule
For Each I In ModItrzP(P)
    Set M = I
    If HasPubMth(Src(M), PubMthn) Then PushI ModNyzPjPubMth, Mdn(M)
Next
#End If
End Function

Function MthTy$(Lin)
MthTy = PfxzAyS(RmvMthMdy(Lin), MthTyAy)
End Function

Sub Z_MthTy()
Dim O$(), L
For Each L In SrczMdn("Fct")
    Push O, MthTy(CStr(L)) & "." & L
Next
BrwAy O
End Sub

Sub Z_MthKd()
Dim A$
Ept = "Property": A = "Property Get": GoSub Tst
Ept = "Property": A = "Property Get":         GoSub Tst
Ept = "Property": A = " Property Get":        GoSub Tst
Ept = "Property": A = "Friend Property Get":  GoSub Tst
Ept = "Property": A = "Friend  Property Get": GoSub Tst
Ept = "":         A = "FriendProperty Get":   GoSub Tst
Exit Sub
Tst:
    Act = MthKd(A)
    C
    Return
End Sub


Sub Z_MthnsetoWiVerb()
MthnsetoWiVerb.Srt.Vc
End Sub

Sub PushNDupDy(ODy(), Dr)
If HasDr(ODy, Dr) Then Exit Sub
PushI ODy, Dr
End Sub

Property Get WiVerbMthNy() As String()
Dim Mthn, I, J&
For Each I In Itr(MthNyV)
    Mthn = I
    If J Mod 100 = 0 Then Debug.Print J
    If HasVerb(Mthn) Then PushI WiVerbMthNy, Mthn
    J = J + 1
Next
End Property

Property Get MthNyoNoVerb() As String()
Dim Mthn, I
For Each I In Itr(MthNyV)
    Mthn = I
    If Not HasVerb(Mthn) Then PushI MthNyoNoVerb, Mthn
Next
End Property

Function HasVerb(Nm) As Boolean
HasVerb = Verb(Nm) <> ""
End Function

Property Get MthnsetoWiVerb() As Aset
Set MthnsetoWiVerb = AsetzAy(WiVerbMthNy)
End Property

Property Get MthnsetoNoVerb() As Aset
Set MthnsetoNoVerb = AsetzAy(MthNyoNoVerb)
End Property

Function MthnsetV() As Aset
Set MthnsetV = AsetzAy(MthNyV)
End Function

Function MthNyzSI(Src$(), MthIxy&()) As String()
Dim Ix
For Each Ix In Itr(MthIxy)
    PushI MthNyzSI, Mthn3zL(Src(Ix)).Nm
Next
End Function

Function MthNyV() As String()
MthNyV = MthNyzV(CVbe)
End Function

Function MthnsetP() As Aset
Set MthnsetP = AsetzAy(MthNyP)
End Function

Function MthNyP() As String()
MthNyP = MthNyzP(CPj)
End Function

Function MthNyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MthNyzP, MthNyzM(C.CodeModule)
Next
End Function

Function MthNyzFb(Fb) As String()
MthNyzFb = MthNyzV(VbezPjf(Fb))
ClsPjf Fb
End Function


Sub Z_MthNyzFb()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
    For Each Fb In AppFbAy
        PushAy O, MthNyzFb(CStr(Fb))
    Next
    Brw O
    Return
X_BrwOne:
    Dim A$(): A = AppFbAy
    Brw MthNyzFb(A(0))
    Return
End Sub

Function MthNy(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB MthNy, Mthn(L)
Next
End Function

Function PubMthNyzM(M As CodeModule) As String()
PubMthNyzM = PubMthNyzS(Src(M))
End Function


Function MthNyzM(M As CodeModule) As String()
MthNyzM = MthNy(Src(M))
End Function

Sub Z_MthnzS()
GoSub Z
Exit Sub
Z:
   B MthNy(SrczP(CPj))
   Return
End Sub

Function MthNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MthNyzV, MthNyzP(P)
Next
End Function

Function MthAsetVbe() As Aset
Set MthAsetVbe = AsetzAy(MthNyV)
End Function

Property Get MthNyzCMd() As String()
MthNyzCMd = MthNyzM(CMd)
End Property

Function HasMth(Src$(), Mthn) As Boolean
HasMth = FstMthIxzN(Src, Mthn) >= 0
End Function

Function HasMthzM(M As CodeModule, Mthn) As Boolean
HasMthzM = HasMth(Src(M), Mthn)
End Function

Function MthnCmlSetVbe() As Aset
Set MthnCmlSetVbe = CmlSetzNy(MthNyV)
End Function

Sub BrwMthnP()
BrwDrs SrtDrs(DoMthnP)
End Sub

Function DoMthnP() As Drs
DoMthnP = SelDrs(DoMthcP, FFoMthn)
End Function

Function DoMthnzV(V As Vbe) As Drs

End Function

Function DoMthnV() As Drs
DoMthnV = DoMthnzV(CVbe)
End Function

Function DoMthnM() As Drs
DoMthnM = DoMthnzM(CMd)
End Function

Function PrpNyzCmp(A As VBComponent) As String()
PrpNyzCmp = Itn(A.Properties)
End Function

Function MthRetTy$(Lin)
Dim A$: A = AftBkt(Lin)
If ShfTerm(A, "As") Then MthRetTy = T1(A)
End Function
