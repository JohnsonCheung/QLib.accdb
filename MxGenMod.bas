Attribute VB_Name = "MxGenMod"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxGenMod."

Sub Cmd_SavGendMod()
Const Lon$ = "T_GenMod"
Dim Lo As ListObject, FxaCell As Range, Fxa$, Pj As VBProject, MdnCell As Range
Dim Mdn$, Md As CodeModule
Dim NewSrc$(), OldL$, NewL$
Dim Ws As Worksheet
Const C$ = "WsGenMod"
:             If Not HasWsCd(C, IsInf:=True) Then Exit Sub          ' Exit=>
     Set Ws = WszCd(C)
:             If Not HasLo(Ws, Lon, IsInf:=True) Then Exit Sub
     Set Lo = Ws.ListObjects(Lon)
:             If Not IsEqCC(Lo, "SrcCd", IsInf:=True) Then Exit Sub ' Exit=>
Set FxaCell = Ws.Range("A1")
        Fxa = FxaCell.Value
        '      If ChkNotHasFfn(Fxa) Then FxaCell.Activate: Exit Sub
              OpnFxa Fxa
     Set Pj = PjzFxa(Fxa)
Set MdnCell = Ws.Range("A2")
        Mdn = MdnCell.Value
:             If ChkMdn(Pj, Mdn) Then MdnCell.Activate: Exit Sub
:             ActWs Ws
     Set Md = MdzP(Pj, Mdn)
     NewSrc = StrColzWsLC(Lo, "T_GenMod", "SrcCd")
:             If Si(NewSrc) = 0 Then MsgBox "No source code is generated", vbCritical: Exit Sub
       OldL = Srcl(Md)
       NewL = JnCrLf(NewSrc)
:             If OldL <> NewL Then MsgBox "The source in WsSrc <> the soruce from Mod", vbCritical: Exit Sub
:             PutAyV NewSrc, A1zLo(Lo)                                                                       ' <== Save
:             MsgBox "SrcCd saved:" & vbCrLf & WrdStszLines(NewL), vbInformation
End Sub

Sub Cmd_LoadSrcCd()
Const C$ = "WsSrcCd"
Const Lon$ = "T_SrcCd"
Dim Ws As Worksheet
Dim Lo As ListObject
Dim FxaCell As Range
Dim Fxa$
Dim Mdn$
Dim Pj As VBProject
Dim MdnCell As Range
Dim Md As CodeModule
Dim SrcCd$()
:             If Not HasWsCd(C, IsInf:=True) Then Exit Sub
     Set Ws = WszCd(C)
:             If Not HasLo(Ws, Lon, IsInf:=True) Then Exit Sub
     Set Lo = Ws.ListObjects(Lon)
:             If Not IsEqCC(Lo, "SrcCd", IsInf:=True) Then Exit Sub
Set FxaCell = Ws.Range("A1")
        Fxa = FxaCell.Value
:          If HasFfn(Fxa) Then FxaCell.Activate: Exit Sub
:              OpnFxa Fxa
     Set Pj = PjzFxa(Fxa)
Set MdnCell = Ws.Range("A2")
        Mdn = MdnCell.Value
:             If ChkMdn(Pj, Mdn) Then MdnCell.Activate: Exit Sub
:             ActWs Ws
     Set Md = MdzP(Pj, Mdn)
      SrcCd = Src(Md)
:             DltLoRow Lo                                                             ' <== Delete
:             PutAyV SrcCd, A1zLo(Lo)                                                 ' <== Load
:             MsgBox "SrcCd loaded:" & vbCrLf & WrdStszLy(SrcCd), vbInformation
End Sub

Sub VdtSrc(ISrc As Drs, IDes As Drs)
Const C$ = "WsGenMod"
Const Lon$ = "T_SrcCd"
:                                        If Not HasWsCd(C, IsInf:=True) Then Exit Sub
Dim Ws      As Worksheet:       Set Ws = WszCd(C)
:                                        If Not HasLo(Ws, Lon, IsInf:=True) Then Exit Sub
Dim Lo      As ListObject:      Set Lo = Ws.ListObjects(Lon)
:                                        If Not IsEqCC(Lo, "SrcCd", IsInf:=True) Then Exit Sub
Dim FxaCell As Range:      Set FxaCell = Ws.Range("A1")
Dim Fxa$:                          Fxa = FxaCell.Value
:                                        If NoFfn(Fxa) Then FxaCell.Activate: Exit Sub
:                                        OpnFxa Fxa
Dim Pj      As VBProject:       Set Pj = PjzFxa(Fxa)
Dim MdnCell As Range:      Set MdnCell = Ws.Range("A2")
Dim Mdn$:                          Mdn = MdnCell.Value
:                                        If ChkMdn(Pj, Mdn) Then MdnCell.Activate: Exit Sub
:                                        ActWs Ws
Dim SrcCd$():                    SrcCd = CdB(ISrc, IDes, SrcCd)
:                                        DltLoRow Lo                                                               ' <== Delete
:                                        PutAyV SrcCd, A1zLo(Lo)                                                   ' <== Load
:                                        MsgBox "SrcCd generated:" & vbCrLf & WrdStszLy(SrcCd), vbInformation
End Sub

Sub Cmd_GenMod(ISrc As Drs, IDes As Drs, InpSrc$(), OupLo As ListObject)
Dim Src$(): Src = CdB(ISrc, IDes, InpSrc)
If HasWs(CWb, "Err") Then MsgBox "There is error", vbCritical: WszWb(CWb, "Err").Activate: Exit Sub
PutCd Src, OupLo
End Sub

Function CdB(Src As Drs, Des As Drs, InpSrc$()) As String()
'Fm Src : V DclSfx Pm Expr !
'Fm Des : V R1 R2 R3       !
'Ret    :                  !  @@
Dim JSrcPm As Drs: JSrcPm = B_JSrcPm(Src) ' ! With Col-Pm added

'== CdRmk OK ===========================================================================================================
Dim RmkVer%:     RmkVer = 2
Dim Rmk1 As Drs:   Rmk1 = SelDrs(JSrcPm, "Key Ret Fss Pm Id StpFor")
Dim Rmk2 As Drs:   Rmk2 = B_Rmk2(RmkVer, Rmk1)
Dim CdRmk$():     CdRmk = AmAddPfx(JnDy(Rmk2.Dy), "                 ' ")

'== CdDes OK ===========================================================================================================
Dim CdDes$(): CdDes = AmAddPfx(JnDy(Des.Dy), " ' ")

'== CdCnst OK ==========================================================================================================
Dim JCnst As Drs:  JCnst = SelDrs(Src, "Key Fss")
Dim CdCnst$():    CdCnst = B_CdCnst(JCnst)

'== CdMain OK ==========================================================================================================
Dim InpPm As Drs:    InpPm = DwEQSel(Src, "StpTy", "Inp", "Key Ret StpTy")
Dim MMPm$:            MMPm = B_MPm(InpPm)
Dim LAs():             LAs = SelDrs(LasRec(Src), "Key Ret StpTy").Dy(0)
Dim MMthn$:          MMthn = LAs(0)
Dim MMRetAs$:      MMRetAs = LAs(1)
Dim LasStpTy$:    LasStpTy = LAs(2)
Dim MMthSorF$:    MMthSorF = B_MMthSorF(LasStpTy)
Dim MMthLin$:      MMthLin = B_MMthLin(MMthSorF, MMthn, MMPm, MMRetAs)

Dim MBdy As Drs: MBdy = DwNeSel(JSrcPm, "StpTy", "Inp", "Key StpTy Stmt Expr Ret Pm BrkNm BrkChr")

Dim MBdy3Col As Drs: MBdy3Col = XMBdy3Col
    If True Then
MBdy3Col = B_MBdy3Col1(MBdy)
    Else
MBdy3Col = B_MBdy3Col(MBdy)
    End If
Dim MMLy$():     MMLy = JnDy(MBdy3Col.Dy)
Dim CdMain$(): CdMain = Sy(MMthLin, CdRmk, CdDes, MMLy, "End " & MMthSorF)

'== CdY ================================================================================================================
Dim CdYLisa  As Drs:  CdYLisa = DwInSel(JSrcPm, "StpTy", SyzSS("Stmt Sub"), "Key StpTy Expr Stmt Ret Pm")
Dim CdYLisb  As Drs:  CdYLisb = DwInSel(JSrcPm, "StpTy", SyzSS("Expr Fun Inp"), "Key StpTy Expr Stmt Ret Pm")
Dim CdY4Cola As Drs: CdY4Cola = B_CdY4Cola(CdYLisa)
Dim CdY4Colb As Drs: CdY4Colb = B_CdY4Colb(CdYLisb)
Dim CdY$():               CdY = Sy(JnDy(CdY4Cola.Dy), Sy(JnDy(CdY4Colb.Dy)))

'== CdZZ OK ============================================================================================================
Dim CdZZLis  As Drs:  CdZZLis = SelDrs(JSrcPm, "Key StpTy Stmt Expr Ret")
Dim CdZZ4Col As Drs: CdZZ4Col = B_CdZZ4Col(CdZZLis)
Dim CdZZ$():             CdZZ = JnDy(CdZZ4Col.Dy)

'== CdB ================================================================================================================
'-- BBMLin OK ----------------------------------------------------------------------------------------------------------
Dim BLis As Drs: BLis = DwInSel(JSrcPm, "StpTy", SyzSS("Sub Fun"), "Key StpTy Ret Pm Fss StpFor")

Dim BCKey$():                  BCKey = SyzDrsC(BLis, "Key")
Dim BCStpTy$():              BCStpTy = SyzDrsC(BLis, "StpTy")
Dim BCRet$():                  BCRet = SyzDrsC(BLis, "Ret")
Dim BCPm$():                    BCPm = SyzDrsC(BLis, "Pm")
Dim BCArgSfx$():            BCArgSfx = B_BCArgSfx(BCRet)
Dim BCCArg As Dictionary: Set BCCArg = DiczAyab(BCKey, BCArgSfx)

Dim BCSorF$():   BCSorF = B_BCSorF(BCStpTy)
Dim BCMthn$():   BCMthn = B_BCMthn(BCKey)
Dim BCTyChr$(): BCTyChr = B_BCTyChr(BCRet, BCStpTy)
Dim BCPmStr$(): BCPmStr = B_BCPmStr(BCPm, BCCArg)
Dim BCRetAs$(): BCRetAs = B_BCRetAs(BCRet, BCStpTy)
Dim BBMLin$():   BBMLin = B_BCMLin(BCSorF, BCMthn, BCTyChr, BCPmStr, BCRetAs)

'-- BBRmkRet -----------------------------------------------------------------------------------------------------------
Dim BCFss$():       BCFss = SyzDrsC(BLis, "Fss")
Dim BCStpFor$(): BCStpFor = SyzDrsC(BLis, "StpFor")
Dim BBRmkRet$(): BBRmkRet = B_BBRmkRet(BCFss, BCStpFor)

'-- BBRmkFm ------------------------------------------------------------------------------------------------------------
Dim BBRmkFm$(): BBRmkFm = B_BBRmkFm(BCPm, BCKey, BCFss, BCStpFor)

'-- BBRmkDes -----------------------------------------------------------------------------------------------------------
Dim BBRmkDes$()

'-- BBCxt OK -----------------------------------------------------------------------------------------------------------
Dim BKey$()
Dim UsedMthn$(): UsedMthn = AmAddPfx(BKey, "B_")

Dim MDic   As Dictionary:   Set MDic = DiMthnqLines(InpSrc, ExlDcl:=True)
Dim MDic1  As Dictionary:  Set MDic1 = B_MDic1(MDic)
Dim MPair  As DicAB:           MPair = DicabzInKy(MDic1, UsedMthn)
Dim MExist As Dictionary: Set MExist = MPair.A
Dim MLines$():                MLines = B_MLines(BCKey, MExist)
Dim BBCxt$():                  BBCxt = B_BBCxt(MLines)

'-- BBExtra OK ---------------------------------------------------------------------------------------------------------
Dim MExtra As Dictionary: Set MExtra = MPair.B
Dim BBExtra$:                BBExtra = JnDblCrLf(SyzItr(SrtDic(MExtra).Items))

'-- BBELin  ------------------------------------------------------------------------------------------------------------
Dim BBELin$(): BBELin = AmAddPfx(BCSorF, "End ")

'-- CdB  ---------------------------------------------------------------------------------------------------------------
CdB = B_CdB(BBMLin, BBRmkFm, BBRmkRet, BBRmkDes, BBCxt, BBELin, BBExtra)

Dim CdOpt$(): CdOpt = B_CdOpt
                CdB = Sy(CdOpt, CdCnst, CdMain, CdB, CdY, CdZZ)
Brw CdB
Stop
End Function

Function B_CdB(MLin$(), RmkFm$(), RmkRet$(), RmkDes$(), Cxt$(), ELin$(), Extra$) As String()
Dim Lines$, O$(), J%
For J = 0 To UB(MLin)
    Erase XX
    X MLin(J)
    X RmkFm(J)
    X RmkRet(J)
'    X RmkDes(J)
    X Cxt(J)
    X ELin(J)
    Lines = JnCrLf(AeEmpEle(XX)) & vbCrLf
    PushI B_CdB, Lines
Next
Erase XX
End Function

Function B_BBRmkFm(Pm$(), Key$(), Fss$(), StpFor$()) As String()
Dim J%, ODy(), IPm, D1 As Dictionary, RmkFm$, D2 As Dictionary
Set D1 = DiczAyab(Key, Fss)
Set D2 = DiczAyab(Key, StpFor)
For J = 0 To UB(Pm)
    Erase ODy
    For Each IPm In Itr(SyzSS(Pm(J)))
        PushI ODy, Array("'Fm :", IPm, D1(IPm), PpdIf(D2(IPm), "! "))
    Next
    RmkFm = JnCrLf(JnDy(ODy))
    PushI B_BBRmkFm, RmkFm
Next
End Function

Function B_BBRmkDes(K$(), CdDes$()) As String()
Dim D As New Dictionary, I
Set D = Dic(RmvFstChrzAy(CdDes), JnSep:=vbCrLf)
For Each I In K
    If D.Exists(I) Then
        PushI B_BBRmkDes, D(I)
    Else
        PushI B_BBRmkDes, ""
    End If
Next
End Function
Function B_BBRmkRet(Fss$(), StpFor$()) As String()
Dim J%
For J = 0 To UB(Fss)
    PushI B_BBRmkRet, "'Ret: " & Fss(J) & PpdIf(StpFor(J), " ! ")
Next
End Function
Function B_MLines(BCKey$(), MExist As Dictionary) As String()
Dim K
For Each K In BCKey
    If MExist.Exists(K) Then
        PushI B_MLines, MExist(K)
    Else
        PushI B_MLines, ""
    End If
Next
End Function

Function B_MDic1(MDic As Dictionary) As Dictionary
Set B_MDic1 = New Dictionary
Dim K
For Each K In MDic.Keys
    B_MDic1.Add BefDot(K), MDic(K)
Next
End Function

Function B_BCArgSfx(Ret$()) As String()
Dim R
For Each R In Ret
    PushI B_BCArgSfx, ArgSfxzRet(R)
Next
End Function

Function B_BCPmStr(Pm$(), DicArgNmToArgSfx As Dictionary) As String()
Dim D As New Dictionary
Set D = DicArgNmToArgSfx
Dim IPm, O$()
For Each IPm In Pm
    Erase O
    Dim Arg$, Sfx$
    Dim P
    For Each P In Itr(SyzSS(IPm))
        Sfx = D(P)
        Arg = P & Sfx
        PushI O, Arg
    Next
    Dim OPm$
    OPm = JnCommaSpc(O)
    PushI B_BCPmStr, OPm
Next
End Function

Function B_BCRetAs(Ret$(), StpTy$()) As String()
Dim J%
For J = 0 To UB(Ret)
    If StpTy(J) = "Sub" Then
        PushI B_BCRetAs, ""
    Else
        PushI B_BCRetAs, RetAs(Ret(J))
    End If
Next
End Function

Function B_BCTyChr(Ret$(), StpTy$()) As String()
Dim J%, O$
For J = 0 To UB(Ret)
    If StpTy(J) = "Fun" Then
        O = TyChrzRet(Ret(J))
    Else
        O = ""
    End If
    PushI B_BCTyChr, O
Next
End Function

Function B_BCMLin(SorF$(), Mthn$(), TyChr$(), Pm$(), RetSfx$()) As String()
Dim J%
For J = 0 To UB(SorF)
    PushI B_BCMLin, FmtQQ("Private ? ??(?)?", SorF(J), Mthn(J), TyChr(J), Pm(J), RetSfx(J))
Next
End Function
Function B_BCSorF(StpTy$()) As String()
Dim T
For Each T In StpTy
    Select Case T
    Case "Sub": PushI B_BCSorF, "Sub"
    Case "Fun": PushI B_BCSorF, "Function"
    Case Else: Stop
    End Select
Next
End Function
Function B_BCMthn(Key$()) As String()
Dim K
For Each K In Key
    PushI B_BCMthn, "B_" & K
Next
End Function

Function B_BBCxt(MLines$()) As String()
Dim Lines
For Each Lines In Itr(MLines)
    PushI B_BBCxt, MthCxt(SplitCrLf(Lines))
Next
End Function

Function B_BBLinesAy(MLin$(), Rmk$(), Cxt$(), ELin$()) As String()
Dim J%
For J = 0 To UB(MLin)
    PushI B_BBLinesAy, JnCrLfAp(MLin(J), Rmk(J), Cxt(J), ELin(J))
Next
End Function

Function B_CdY4Cola(CdYLisa As Drs) As Drs
B_CdY4Cola = B_CdY4Colb(CdYLisa)
End Function

Function B_CdY4Colb(CdYLisb As Drs) As Drs
'CdYLis  Key StpTy Expr Stmt Ret Pm
'CdY4Col Fun YEq Callg End
Dim ODy(), YEq$, Dr, Key$, TyChr$, RetSfx$, Ret$, Fun$, Callg$, StpTy$, Expr$, Pm$, Stmt$, MthEndLin$
For Each Dr In Itr(CdYLisb.Dy)
    Key = Dr(0)
    StpTy = Dr(1)
    Expr = Dr(2)
    Stmt = Dr(3)
    Ret = Dr(4)
    Pm = Dr(5)
    Pm = JnCommaSpc(AmAddPfx(SyzSS(Pm), "Y_"))
    RetSfx = RetAs(Ret)
    TyChrzRet (Ret)
    Select Case StpTy
    Case "Inp", "Fun", "Expr"
        Fun = FmtQQ("Function Y_??()?:", Key, TyChr, RetSfx)
        MthEndLin = "End Function"
    Case "Stmt", "Sub"
        Fun = FmtQQ("Sub Y_??()?:", Key, TyChr, RetSfx)
        MthEndLin = "End Sub"
    Case Else
        Thw CSub, "StpTy Err:", "StpTy", StpTy
    End Select
    Select Case StpTy
    Case "Inp"
        YEq = ""
        Callg = ""
    Case "Fun"
        YEq = FmtQQ("Y_? =", Key)
        Callg = FmtQQ("B_?(?):", Key, Pm)
    Case "Expr"
        YEq = FmtQQ("Y_? =", Key)
        Callg = Expr & ":"
    Case "Sub"
        YEq = ""
        Callg = FmtQQ("B_? ?:", Key, Pm)
    Case "Stmt"
        YEq = FmtQQ("Y_? =", Key)
        Callg = Stmt & ":"
    Case Else
        Thw CSub, "StpTy er should Be Inp|Expr|Fun", "StpTy", StpTy
    End Select
    PushI ODy, Array(Fun, YEq, Callg, MthEndLin)
Next
ODy = AlignRzDyC(ODy, 1)
B_CdY4Colb = DrszFF("Fun YEq Callg End", ODy)
End Function

Function B_CdZZ4Col(CdZZLis As Drs) As Drs
'CdZZLis  Key StpTy Stmt Expr Ret
'CdZZ4Col Fun Brwg Y End
Dim ODy(), Dr, Key$, TyChr$, RetSfx$, Ret$, Fun$, StpTy$, Expr$, Pm$, Stmt$, Y$, Brwg$
For Each Dr In Itr(CdZZLis.Dy)
    Key = Dr(0)
    StpTy = Dr(1)
    Stmt = Dr(2)
    Expr = Dr(3)
    Ret = Dr(4)
    Fun = FmtQQ("Private Sub Z_?():", Key)
    Select Case StpTy
    Case "Inp", "Fun", "Expr"
        Y = "Y_" & Key
        'Brwg = BrwrzRet(Ret)
    Case "Stmt", "Sub"
        Brwg = ""
        Y = "Y_" & Key
    Case Else
        Thw CSub, "StpTy er should Be Inp|Expr|Fun", "StpTy", StpTy
    End Select
    PushI ODy, Array(Fun, Brwg, Y & ":", "End Sub")
Next
ODy = AlignRzDyC(ODy, 1)
B_CdZZ4Col = DrszFF("Fun Brwg Y End", ODy)
End Function

Function Y_ISrc() As Drs
Erase XX
X ""
Y_ISrc = DrszFmtg(XX)
Erase XX
End Function

Function B_JSrcPm(ISrc As Drs) As Drs
'Fm : ISrc ! From Xls Lo T_Src
'Fm : CC   ! = Key Pfx Id StpTy StpFor Ret Fss Ret Stmt Expr BrkNm BrkChr Fm1..5
'Ret: Key Pfx Id StpTy StpFor Ret Fss Ret Stmt Expr BrkNm BrkChr Pm
Const CC$ = "Key Pfx Id StpTy StpFor Ret Fss Ret Stmt Expr BrkNm BrkChr Fm1 Fm2 Fm3 Fm4 Fm5"
Dim Dr, I1%, I2%, I3%, I4%, I5%, ODy()
Dim O As Drs: O = SelDrs(ISrc, CC)
AsgIx ISrc, "Fm1 Fm2 Fm3 Fm4 Fm5", I1, I2, I3, I4, I5
For Each Dr In Itr(ISrc.Dy)
    PushI Dr, JnSpc(SyNB(Dr(I1), Dr(I2), Dr(I3), Dr(I4), Dr(I5)))
    PushI ODy, Dr
Next
Stop '
'B_JSrcPm = DeCC(Drs(Sy(ISrc.Fny, "Pm"), ODy), "Fm*")
End Function
Function B_MMCDcl(Key$(), StpTy$(), Ret$()) As String()
Dim J%, O$, ArgSfx$
For J = 0 To UB(Key)
    Select Case StpTy(J)
    Case "Sub", "Stmt": O = ""
    Case "Fun", "Expr"
        ArgSfx = ArgSfxzRet(Ret(J))
        O = FmtQQ("Dim B_??:", Key(J), ArgSfx)
    Case Else: Thw CSub, "Invalid StpTy.  Should be Sub|Fun|Stmt|Expr", "[StpTy with err] RowIx MBdy", StpTy(J), J, FmtAyab(Key, StpTy)
    End Select
    PushI B_MMCDcl, O
Next
End Function
Function B_MMCLHS(Key$(), StpTy$()) As String()
Dim J%
For J = 0 To UB(Key)
    Select Case StpTy(J)
    Case "Sub", "Stmt": PushI B_MMCLHS, ""
    Case "Fun", "Expr": PushI B_MMCLHS, "B_" & Key(J) & " ="
    Case Else:          Thw CSub, "Invalid StpTy.  Should be Sub|Fun|Stmt|Expr", "[StpTy with err] RowIx MBdy", StpTy(J), J, FmtAyab(Key, StpTy)
    End Select
Next
End Function
Function B_MMCCallg(Key$(), StpTy$(), Pm$(), Stmt$(), Expr$()) As String()
Dim J%, O$
For J = 0 To UB(StpTy(J))
    Select Case StpTy(J)
    Case "Stmt":               O = Stmt(J)
    Case "Sub":                O = Key(J) & " " & Pm(J)
    Case "Fun" And Pm(J) = "": O = Key(J)
    Case "Fun":                O = Key(J) & QteBkt(Pm(J))
    Case "Expr":               O = Expr(J)
    Case Else: Thw CSub, "Invalid StpTy.  Should be Sub|Fun|Stmt|Expr", "[StpTy with err] RowIx MBdy", StpTy(J), J, FmtAyab(Key, StpTy)
    End Select
    Push B_MMCCallg, O
Next
End Function
Function B_MBdy3Col1(MBdy As Drs) As Drs
'MBldPm   Key StpTy Stmt Expr Ret Pm
'MBdy3Col Dcl LHS Calling
Dim Key$(), StpTy$(), Ret$(), Pm$(), Stmt$(), Expr$()
ColApzDrs MBdy, "Key StpTy Ret Pm Stmt Expr", Key, StpTy, Ret, Pm, Stmt, Expr
Dim J%, U%
U = UB(MBdy.Dy)
Dim Dcl$()
    ReDim Dcl(U)
    Dim ArgSfx$
    For J = 0 To U
        Select Case StpTy(J)
        Case "Sub", "Stmt"
        Case "Fun", "Expr"
            ArgSfx = ArgSfxzRet(Ret(J))
            Dcl(J) = FmtQQ("Dim B_??:", Key(J), ArgSfx)
        Case Else: Thw CSub, "Invalid StpTy.  Should be Sub|Fun|Stmt|Expr", "[StpTy with err] RowIx MBdy", StpTy, J, FmtDrszNoRdu(MBdy)
        End Select
    Next
Dim LHS$()
    ReDim LHS(U)
    For J = 0 To U
        Select Case StpTy(J)
        Case "Sub", "Stmt"
        Case "Fun", "Expr": LHS(J) = "B_" & Key(J) & " ="
        Case Else: Stop
        End Select
    Next
Dim Callg$()
    ReDim Callg(U)
    Dim T$, P$
    For J = 0 To U
        T = StpTy(J)
        P = JnCommaSpc(SyzSS(Pm(J)))
        Select Case True
        Case T = "Stmt":               Callg(J) = Stmt(J)
        Case T = "Sub":                Callg(J) = Key(J) & " " & P
        Case T = "Fun" And Pm(J) = "": Callg(J) = Key(J)
        Case T = "Fun":                Callg(J) = Key(J) & QteBkt(P)
        Case T = "Expr":               Callg(J) = Expr(J)
        Case Else: Stop
        End Select
    Next

Dim ODy()
    LHS = AlignRzAy(LHS)
    For J = 0 To U
        PushI ODy, Array(Dcl(J), LHS(J), Callg(J))
    Next
B_MBdy3Col1 = DrszFF("Dcl LHS Callg", ODy)
End Function

Function B_MBdy3Col(MBdy As Drs) As Drs
'         0   1     2    3    4   5
'MBldPm   Key StpTy Stmt Expr Ret Pm
'MBdy3Col Dcl LHS Calling
Dim ODy()
Dim Dr, DclSfx$, Dcl$, Key$, Ix%, Pm$, Expr$, Stmt$, StpTy$, LHS$, Callg$, Ret$
Dim KeyAy$(): KeyAy = StrColzDy(MBdy.Dy, 0)
For Each Dr In MBdy.Dy
    Key = Dr(0)
    Pm = Dr(5)
    Expr = Dr(3)
    Stmt = Dr(2)
    StpTy = Dr(1)
    Select Case StpTy
    Case "Sub", "Stmt"
        Dcl = ""
        LHS = ""
        If StpTy = "Stmt" Then
            Callg = Stmt
        Else
            Callg = Key & " " & Pm
        End If
    Case "Fun", "Expr"
        Ret = Dr(4)
        DclSfx = ArgSfxzRet(Ret)
        Dcl = FmtQQ("Dim B_??:", Key, DclSfx)
        LHS = "B_" & Key & " ="
        If StpTy = "Fun" Then
            If Pm = "" Then
                Callg = Key
            Else
                Callg = Key & QteBkt(Pm)
            End If
        Else
            Callg = Expr
        End If
    Case Else: Thw CSub, "Invalid StpTy.  Should be Sub|Fun|Stmt|Expr", "[StpTy with err] RowIx MBdy", StpTy, Ix, FmtDrszNoRdu(MBdy)
    End Select
    PushI ODy, Array(Dcl, LHS, Callg)
    Ix = Ix + 1
Next
ODy = AlignRzDyC(ODy, 1)
B_MBdy3Col = DrszFF("Dcl LHS Callg", ODy)
End Function

Function B_MMthSorF$(LasStpTy$)
Select Case LasStpTy
Case "Fun": B_MMthSorF = "Function"
Case "Sub": B_MMthSorF = "Sub"
Case Else: Thw CSub, "The StpTy of Las Record of WsSrc must be Either Fun or Sub", "LasStpTy", LasStpTy
End Select
End Function

Function B_Rmk2(TRmkVer%, JRmk As Drs) As Drs
'JRmk Key Ret Fss Pm Id
'     0   1   2   3  4
Dim ODy(), JRmkDy()
JRmkDy = JRmk.Dy
Dim Ret$, Fss$, C$, Id$
Dim Dr
Select Case TRmkVer
Case 1
    For Each Dr In Itr(JRmkDy)
        Ret = Dr(1)
        Fss = Dr(2)
        If Ret = "Drs" Then
            C = Fss
        Else
            C = Ret
        End If
        PushI ODy, Array(Dr(0), C, "| " & Dr(3))
    Next
    B_Rmk2 = DrszFF("Rmk1 2 3", ODy)
Case 2
    Dim Using$, Pm$
    Dim Dic As Dictionary 'Key-To-Id
    Set Dic = DiczDyCC(JRmkDy, 0, 4)
    For Each Dr In Itr(JRmkDy)
        Ret = Dr(1)
        Fss = Dr(2)
        Pm = Dr(3)
        Id = Dr(4)
        If Ret = "Drs" Then
            C = Fss
        Else
            C = Ret
        End If
        Using = JnSpc(VyzDicKK(Dic, Pm))
        PushI ODy, Array(Dr(0), Id, C, "| " & Using)
    Next
    B_Rmk2 = DrszFF("Rmk1 2 3 4", ODy)
Case Else: Thw CSub, "TRmkVer should 1 or 2", "TRmkVer", TRmkVer
End Select
End Function

Function B_MPm$(InpPm As Drs)
'InpPm Key Ret
Dim O$(), Dr, Sfx$, Ret$
For Each Dr In Itr(InpPm.Dy)
    PushI O, Dr(0) & ArgSfxzRet(Dr(1))
Next
B_MPm = JnCommaSpc(O)
End Function

Sub Z_CdB()
Dim Cd$(), ISrc As Drs, IDes As Drs, SrcCd$()
GoSub Z
Exit Sub
Z:
    Brw CdB(ISrc, IDes, SrcCd)
    Return
End Sub

Function B_CdOpt() As String()
PushI B_CdOpt, "Option Explicit"
PushI B_CdOpt, "Option Compare Text"
End Function

Function B_CdCnst(JCnst As Drs) As String()
'Fm : JCnst Key Fss
'Ret: ! SrcCd for module constant
Dim Dr
For Each Dr In Itr(JCnst.Dy)
    PushI B_CdCnst, FmtQQ("Const ?_$ = ""?""", Dr(0), Dr(1))
Next
End Function

Function B_MPmLin$(MainPm As Drs)
'Fm  : MainPm (MainDD) : Nm TyChr AsTy !
'Ret : !
Dim O$(), Arg$, Nm$, TyChr$, AsTy$, Dr, INm%, ITyChr%, IAsTy%
AsgIx MainPm, "Nm TyChr AsTy", INm, ITyChr, IAsTy
For Each Dr In Itr(MainPm.Dy)
    Nm = Dr(0)
    TyChr = Dr(1)
    AsTy = Dr(2)
    Arg = Nm & TyChr & PpdIf(AsTy, " As ")
    PushI O, Arg
Next
B_MPmLin = JnCommaSpc(O)
End Function

Function B_MMthLin$(MthTy$, Mthn$, Pm$, Ret$)
'MthTy is Either Function or Sub
Dim TyChr$, RetSfx$
If IsTyChr(Ret) Then
    TyChr = Ret
Else
    RetSfx = Ret
End If
B_MMthLin = FmtQQ("? ??(?)?", MthTy, Mthn, TyChr, Pm, RetSfx)
End Function


Function XMBdy3Col() As Drs
'Insp "QIde_Gen_GenMod.XMBdy3Col", "Inspect", "Oup(XMBdy3Col) ", FmtDrszNoRdu(XMBdy3Col), : Stop
End Function

