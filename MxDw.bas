Attribute VB_Name = "MxDw"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDw."

Function AddColzFst(D As Drs, Gpcc$) As Drs
'Fm D    : ..@Gpcc.. ! a drs with col-@Gpcc
'Fm Gpcc :           ! col-@Gpcc in @D have dup.
'Ret     : @D Fst    ! a drs of col-Fst add to @D at end.  col-Fst is bool value.  TRUE when if it fst rec of a gp
'                    ! and rst of rec of the gp to FALSE
Dim O As Drs: O = AddCol(D, "Fst", False) ' Add col-Fst with val all FALSE
If NoReczDrs(D) Then AddColzFst = O: Exit Function
Dim GDy(): GDy = SelDrs(D, Gpcc).Dy  ' Dy with Gp-col only.
Dim R(): R = GpRxy(GDy)                 ' Gp the @GDy into `GpRxy`
Dim Cix&: Cix = UB(O.Dy(0))             ' Las col Ix aft adding col-Fst
Dim Rxy: For Each Rxy In R               ' for each gp, get the Row-ixy (pointing to @D.Dy)
    Dim Rix&: Rix = Rxy(0)               ' Rix is Row-ix pointing one of @D.Dy which is the fst rec of a gp
    O.Dy(Rix)(Cix) = True
Next
AddColzFst = O
End Function

Sub AsgColonFF(ColonFF$, OFnyA$(), OFnyB$())
Erase OFnyA, OFnyB
Dim F: For Each F In SyzSS(ColonFF)
    With BrkBoth(F, ":")
        PushI OFnyA, .S1
        PushI OFnyB, .S2
    End With
Next
End Sub

Function DeCeqC(A As Drs, CC$) As Drs
Dim Dr, C1&, C2&
AsgIx A, CC, C1, C2
For Each Dr In Itr(A.Dy)
    If Dr(C1) <> Dr(C2) Then
        PushI DeCeqC.Dy, Dr
    End If
Next
DeCeqC.Fny = A.Fny
End Function

Function DeDup(A As Drs) As Drs
DeDup = DeDupzFF(A, JnSpc(A.Fny))
End Function

Function DeDupzFF(A As Drs, FF$) As Drs
Dim Rxy&(): Rxy = F_SubRxy_ByDupFF(A, FF)
DeDupzFF = DeRxy(A, Rxy)
End Function

Function DwCeqC(A As Drs, CC$) As Drs
Dim Dr, C1&, C2&
AsgIx A, CC, C1, C2
For Each Dr In Itr(A.Dy)
    If Dr(C1) = Dr(C2) Then
        PushI DwCeqC.Dy, Dr
    End If
Next
DwCeqC.Fny = A.Fny
End Function

Function DwCneC(A As Drs, CC$) As Drs
DwCneC = DeCeqC(A, CC)
End Function

Function DyoSelAlwE(Dy(), Ixy&()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DyoSelAlwE, AwIxyAlwE(Dr, Ixy)
Next
End Function

Function DywDup(Dy()) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim Dr
For Each Dr In GRxyzCyCnt(Dy)
    If Dr(0) > 1 Then
        PushI DywDup, AeFstEle(Dr)
    End If
Next
End Function

Function DywKey(Dy(), KeyIxy&(), Key()) As Variant()
'Ret : SubSet-of-row of @Dy for each row has val of %CurKey = @Key
Dim Dr: For Each Dr In Itr(Dy)
    Dim CurK: CurK = AwIxy(Dr, KeyIxy)
    If IsEqAy(CurK, Key) Then         '<- If %CurKey = @Key, select it.
        PushI DywKey, Dr
    End If
Next
End Function

Function DywKeySel(Dy(), KeyIxy&(), Key(), SelIxy&()) As Variant()
DywKeySel = SelDy(DywKey(Dy, KeyIxy, Key), SelIxy)
End Function

Function ExpandFF(FF$, Fny$()) As String() '
ExpandFF = ExpandLikAy(TermAy(FF), Fny)
End Function

Function ExpandLikAy(LikAy$(), Ay$()) As String() 'Put each expanded-ele in likAy to return a return ay. _
Expanded-ele means either the ele itself if there is no ele in Ay is like the `ele` _
                   or     the lik elements in Ay with the given `ele`
Dim Lik
For Each Lik In LikAy
    Dim A$()
    A = AwLik(Ay, Lik)
    If Si(A) = 0 Then
        PushI ExpandLikAy, Lik
    Else
        PushIAy ExpandLikAy, A
    End If
Next
End Function

Function FnyAzJn(Jn$) As String()
Dim J
For Each J In SyzSS(Jn)
    PushI FnyAzJn, BefOrAll(J, ":")
Next
End Function

Function FnyBzJn(Jn$) As String()
Dim J
For Each J In SyzSS(Jn)
    PushI FnyBzJn, AftOrAll(J, ":")
Next
End Function

Function FnyzSelAtEnd(Fny$(), AtEndFny$())
FnyzSelAtEnd = AddSy(SyMinus(Fny, AtEndFny), AtEndFny)
End Function

Function FnyzSelInFront(Fny$(), InFrontFny$())
Dim InFront$(): InFront = AyIntersect(InFrontFny, Fny)
FnyzSelInFront = AddSy(InFront, SyMinus(Fny, InFront))
End Function

Function GpRxy(Dy()) As Variant()
'Fm Dy : all col in @Dy will be used to gp
'Ret    : N-Gp of Rxy (Rec-Ix-Ay) pointing to @Dy.  That means each gp contain a Rxy.
'         and each ele in each Rxy is a Rix pointing a dist rec of @Dy
Dim K(), Dr, O(), Rix&: For Each Dr In Itr(Dy)
    Dim Gix&: Gix = IxzDyDr(K, Dr)
    If Gix = -1 Then
        Dim Rxy&(): ReDim Rxy(0)
        Rxy(0) = Rix
        PushI O, Rxy      '<== Put Rix to Oup-O
        PushI K, Dr       '<-- Put Dr to K
    Else
        PushI O(Gix), Rix '<== Put Rix to Oup-O
    End If
    Rix = Rix + 1
Next
GpRxy = O
End Function

Function GRxyzCyCnt(Dy()) As Variant()
#If True Then
    GRxyzCyCnt = GRxyzCyCntzSlow(Dy)
#Else
    GRxyzCyCnt = GRxyzCyCntzQuick(Dy)
#End If
End Function

Function GRxyzCyCntzQuick(Dy()) As Variant()
End Function

Function GRxyzCyCntzSlow(Dy()) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim OKeyDy(), OCnt&(), Dr
    Dim LasIx&: LasIx = Si(Dy(0))
    Dim J&
    For Each Dr In Dy
        If J Mod 500 = 0 Then Debug.Print "GRxyzCyCntzSlow"
        If J Mod 50 = 0 Then Debug.Print J;
        J = J + 1
        With IxOptzDyDr(OKeyDy, Dr)
            Select Case .Som
            Case True: OCnt(.Lng) = OCnt(.Lng) + 1
            Case Else: PushI OKeyDy, Dr: PushI OCnt, 1
            End Select
        End With
    Next
    If Si(OKeyDy) <> Si(OCnt) Then Thw CSub, "Si Diff", "OKeyDy-Si OCnt-Si", Si(OKeyDy), Si(OCnt)
For J = 0 To UB(OCnt)
    PushI GRxyzCyCntzSlow, AddAy(Array(OCnt(J)), OKeyDy(J)) '<===========
Next
End Function

Function InsCol(A As Drs, C$, V) As Drs
InsCol = InsColzFront(A, C, V)
End Function

Function InsColzDrsC3(A As Drs, CCC$, V1, V2, V3) As Drs
InsColzDrsC3 = DwInsFF(A, CCC, InsColzDyV3(A.Dy, V1, V2, V3))
End Function

Function InsColzDrsCC(A As Drs, CC$, V1, V2) As Drs
InsColzDrsCC = DwInsFF(A, CC, InsColzDyV2(A.Dy, V1, V2))
End Function

Function InsColzDyBef(Dy(), V) As Variant()
InsColzDyBef = InsColzDyVyBef(Dy, Av(V))
End Function

Function InsColzDyVyBef(Dy(), Vy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI InsColzDyVyBef, AddAy(Vy, Dr)
Next
End Function

Function InsColzFront(A As Drs, C$, V) As Drs
InsColzFront = DwInsFF(A, C, InsColzDyBef(A.Dy, V))
End Function

Function IxOptzDyDr(Dy(), Dr) As LngOpt
Dim IDr, Ix&
For Each IDr In Itr(Dy)
    If IsEqAy(IDr, Dr) Then IxOptzDyDr = SomLng(Ix): Exit Function
    Ix = Ix + 1
Next
End Function

Function JnDrs(A As Drs, b As Drs, Jn$, Add$, Optional IsLeftJn As Boolean, Optional AnyFld$) As Drs
'Fm A        : ..@Jn-LHS..              ! It is a drs with col-@Jn-LHS.
'Fm B        : ..@Jn-RHS..@Add-RHS      ! It is a drs with col-@Jn-RHS & col-@Add-RHS.
'Fm Jn       : :SS-of-:ColonTerm        ! It is :SS-of-:ColTerm. :ColTerm: is a :Term with 1-or-0 [:]. :Term: is a fm :TLin: or :Termss:  LHS of [:] is for @A and RHS of [:] is for @B
'                                       ! It is used to jn @A & @B
'Fm Add      : SS-of-ColonStr-Fld-@B    ! What col in @B to be added to @A.  It may use new name, if it has colon.
'Fm IsLeftJn :                          ! Is it left join, otherwise, it is inner join
'Fm AnyFld   : Fldn                     ! It is optional fld to be add to rslt drs stating if any rec in @B according to @Jn.
'                                       ! It is vdt only when IsLeftJn=True.
'                                       ! It has bool value.  It will be TRUE if @B has jn rec else FALSE.
'Ret         : ..@A..@Add-RHS..@AnyFld ! It has all fld from @A and @Add-RHS-fld and optional @AnyFld.
'                                       ! If @IsLeftJn, it will have at least same rec as @A, and may have if there is dup rec in @B accord to @Jn fld.
'                                       ! If not @IsLeftJn, only those records fnd in both @A & @B
Dim JnFnyA$(), JnFnyB$()
Dim AddFnyFm$(), AddFnyAs$()
    AsgColonFF Jn, JnFnyA, JnFnyB
    AsgColonFF Add, AddFnyFm, AddFnyAs
    
Dim AddIxy&(): AddIxy = IxyzSubAy(b.Fny, AddFnyFm, ThwNFnd:=True)
Dim BJnIxy&(): BJnIxy = IxyzSubAy(b.Fny, JnFnyB, ThwNFnd:=True)
Dim AJnIxy&(): AJnIxy = IxyzSubAy(A.Fny, JnFnyA, ThwNFnd:=True)

Dim Emp() ' it is for LeftJn and for those rec when @B has no rec joined.  It is for @Add-fld & @AnyFld.
          ' It has sam ele as @Add.  1 more fld is @AnyFld<>""
    If IsLeftJn Then
        ReDim Emp(UB(AddFnyFm))
        If AnyFld <> "" Then PushI Emp, False
    End If
Dim ODy()                       ' Bld %ODy for each %ADr, that mean fld-Add & fld-Any
    Dim Adr: For Each Adr In Itr(A.Dy)
        Dim JnVy():            JnVy = AwIxy(Adr, AJnIxy)                     'JnFld-Vy-Fm-@A
        Dim Bdy():            Bdy = DywKeySel(b.Dy, BJnIxy, JnVy, AddIxy) '@B-Dy-joined
        Dim NoRec As Boolean: NoRec = Si(Bdy) = 0                           'no rec joined
            
        Select Case True
        Case NoRec And IsLeftJn: PushI ODy, AddAy(Adr, Emp) '<== ODy, Only for NoRec & LeftJn
        Case NoRec
        Case Else
            '
            Dim Bdr: For Each Bdr In Bdy
                If AnyFld <> "" Then
                    Push Bdr, True
                End If
                PushI ODy, AddAy(Adr, Bdr) '<== ODy, for each %BDr in %BDy, push to %ODy
            Next
        End Select
    Next Adr

Dim O As Drs: O = Drs(SyNB(A.Fny, AddFnyAs, AnyFld), ODy)
JnDrs = O

If False Then
    Erase XX
    XBox "Debug JnDrs"
    X "A-Fny  : " & Termss(A.Fny)
    X "B-Fny  : " & Termss(b.Fny)
    X "Jn     : " & Jn
    X "Add    : " & Add
    X "IsLefJn: " & IsLeftJn
    X "AnyFld : [" & AnyFld & "]"
    X "O-Fny  : " & Termss(O.Fny)
    X "More ..: A-Drs B-Drs Rslt"
    X LyzNmDrs("A-Drs  : ", A)
    X LyzNmDrs("B-Drs  : ", b)
    X LyzNmDrs("Rslt   : ", O)
    Brw XX
    Erase XX
    Stop
End If
End Function

Function LDrszJn(A As Drs, b As Drs, Jn$, Add$, Optional AnyFld$) As Drs
LDrszJn = JnDrs(A, b, Jn, Add, IsLeftJn:=True, AnyFld:=AnyFld)
End Function

Function F_SubRxy_ByDupFF(A As Drs, FF$) As Long()
Dim Fny$(): Fny = TermAy(FF)
If Si(Fny) = 1 Then
    F_SubRxy_ByDupFF = IxyzDup(ColzDrs(A, Fny(0)))
    Exit Function
End If
Dim ColIxy&(): ColIxy = Ixy(A.Fny, Fny)
Dim Dy(): Dy = SelDy(A.Dy, ColIxy)
F_SubRxy_ByDupFF = F_SubRxy_ByDupFFDy(Dy)
End Function

Function F_SubRxy_ByDupFFDy(Dy()) As Long()
Dim DupD(): DupD = DywDup(Dy)
Dim Dr, Ix&, O&()
For Each Dr In Dy
    If HasDr(DupD, Dr) Then PushI O, Ix
    Ix = Ix + 1
Next
If Si(O) < Si(DupD) * 2 Then Stop
F_SubRxy_ByDupFFDy = O
End Function

Function F_SubRxy_ByDupFFDyColIx(Dy(), ColIx&) As Long()
Dim D As New Dictionary, FstIx&, V, O As New Rel, Ix&, I
For Ix = 0 To UB(Dy)
    V = Dy(Ix)(ColIx)
    If D.Exists(V) Then
        O.PushParChd V, D(V)
        O.PushParChd V, Ix
    Else
        D.Add V, Ix
    End If
Next
For Each I In O.SetOfPar.Itms
    PushIAy F_SubRxy_ByDupFFDyColIx, O.ParChd(I).Av
Next
End Function
Function SelDistFny(D As Drs, Fny$()) As Drs
With GpCntFny(D, Fny)
    SelDistFny = Drs(Fny, .Gp)
End With
End Function
Function SelDistAllCol(D As Drs) As Drs
With GpCntAllCol(D)
    SelDistAllCol = Drs(D.Fny, .Gp)
End With
End Function

Function SelDist(D As Drs, FF$) As Drs
With GpCnt(D, FF)
    SelDist = DrszFF(FF, .Gp)
End With
End Function

Function SelDistCnt(D As Drs, FF$) As Drs
'@D : ..{Gpcc}    ! it has columns-Gpcc
'Ret   : {Gpcc} Cnt  ! each @Gpcc is unique.  Cnt is rec cnt of such gp
Dim Gp(), Cnt&()
    With GpCnt(D, FF)
        Gp = .Gp
        Cnt = .Cnt
    End With
Dim ODy()
    Dim J&, Dr: For Each Dr In Itr(Gp)
        Push Dr, Cnt(J)
        PushI Gp, Dr
        J = J + 1
    Next
Dim Fny$(): Fny = AddEleS(D.Fny, "Cnt")
SelDistCnt = Drs(Fny, ODy)
End Function

Function SelDrs(A As Drs, FF$) As Drs
SelDrs = SelDrsFny(A, SyzSS(FF))
End Function

Function SelDrsAlwE(A As Drs, FF$) As Drs
SelDrsAlwE = SelDrsAlwEzFny(A, SyzSS(FF))
End Function

Function SelDrsAlwEzFny(A As Drs, Fny$()) As Drs
If IsEqAy(A.Fny, Fny) Then SelDrsAlwEzFny = A: Exit Function
SelDrsAlwEzFny = Drs(Fny, DyoSelAlwE(A.Dy, IxyzAlwE(A.Fny, Fny)))
End Function

Function SelDrsAs(A As Drs, AsFF$) As Drs
Dim Fa$(), Fb$(): AsgColonFF AsFF, Fa, Fb
SelDrsAs = Drs(Fb, SelDrsFny(A, Fa).Dy)
End Function

Function SelDrsAtEnd(D As Drs, AtEndFF$) As Drs
Dim NewFny$(): NewFny = FnyzSelAtEnd(D.Fny, SyzSS(AtEndFF))
SelDrsAtEnd = SelDrsFny(D, NewFny)
End Function

Function SelDrsInFront(D As Drs, InFrontFF$) As Drs
Dim NewFny$(): NewFny = FnyzSelInFront(D.Fny, SyzSS(InFrontFF))
SelDrsInFront = SelDrsFny(D, NewFny)
End Function

Function SelDrsExlCC(A As Drs, ExlCCLik$) As Drs
Dim LikC
For Each LikC In SyzSS(ExlCCLik)
'    AyMinus(
Next
End Function

Function SelDrsFny(A As Drs, Fny$()) As Drs
ThwNotSuperAy A.Fny, Fny
Dim I&(): I = Ixy(A.Fny, Fny)
SelDrsFny = Drs(Fny, SelDy(A.Dy, I))
End Function

Function SelDt(A As Dt, FF$) As Dt
SelDt = DtzDrs(SelDrs(DrszDt(A), FF), A.DtNm)
End Function

Function UpdC(A As Drs, C$, V) As Drs
Dim I&: I = IxzAy(A.Fny, C)
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    Dr(I) = V
    PushI Dy, Dr
Next
UpdC = Drs(A.Fny, Dy)
End Function

Function UpdCC(A As Drs, CC$, V1, V2) As Drs
Dim I1&, I2&: AsgIx A, CC, I1, I2
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    Dr(I1) = V1
    Dr(I2) = V2
    PushI Dy, Dr
Next
UpdCC = Drs(A.Fny, Dy)
End Function

Function UpdDrs(A As Drs, b As Drs, Jn$, Upd$, IsLefJn As Boolean) As Drs
'@A  : ..@Jn-LHS..@Upd-LHS.. ! to be updated
'@B  : ..@Jn-RHS..@Upd-RHS.. ! used to update @A.@Upd-LHS
'@Jn : :SS-JnTerm            ! :JnTerm is :ColonTerm.  LHS is @A-fld and RHS is @B-fld
'Fm Upd : :Upd-UpdTerm          ! :UpdTer: is :ColTerm.  LHS is @A-fld and RHS is @B-fld
'Ret    : sam as @A             ! new Drs from @A with @A.@Upd-LHS updated from @B.@Upd-RHS. @@
Dim C As Dictionary: Set C = DiczDrsCC(b)
Dim O As Drs
    O.Fny = A.Fny
    Dim Dr, K
    For Each Dr In A.Dy
        K = Dr(0)
        If C.Exists(K) Then
            Dr(0) = C(K)
        End If
        PushI O.Dy, Dr
    Next
UpdDrs = O
'BrwDrs3 A, B, O, NN:="A B O", Tit:= _
Stop
End Function


Sub Z_DwDup()
Dim A As Drs, FF$, Act As Drs
GoSub T0
Exit Sub
T0:
    A = DrszFF("A B C", Av(Av(1, 2, "xxx"), Av(1, 2, "eyey"), Av(1, 2), Av(1), Av(Empty, 2)))
    FF = "A B"
    GoTo Tst
Tst:
    Act = DwDup(A, FF)
    DmpDrs Act
    Return
End Sub

Sub Z_F_SubRxy_ByDupFFDyColIx()
Dim Dy(), ColIx&, Act&(), Ept&()
GoSub T0
Exit Sub
T0:
    ColIx = 0
    Dy = Array(Array(1, 2, 3, 4), Array(1, 2, 3), Array(2, 4, 3))
    Ept = LngAy(0, 1)
    GoTo Tst
Tst:
    Act = F_SubRxy_ByDupFFDyColIx(Dy, ColIx)
    If Not IsEqAy(Act, Ept) Then Stop
    C
    Return
End Sub

Sub Z_SelDist()
'BrwDrs SelDistCnt(DoPubFun, "Mdn Ty")
End Sub
Function DePatn(A As Drs, C$, ExlPatn$) As Drs
If ExlPatn = "" Then DePatn = A: Exit Function
Dim ODy()
Dim R As RegExp: Set R = Rx(ExlPatn)
Dim Ix%: Ix = IxzAy(A.Fny, C)
Dim Dr: For Each Dr In Itr(A.Dy)
    Dim V: V = Dr(Ix)
    If Not R.Test(V) Then
        PushI ODy, Dr
    End If
Next
DePatn = Drs(A.Fny, ODy)
End Function

