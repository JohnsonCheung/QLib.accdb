Attribute VB_Name = "MxGenErMsg"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxGenErMsg."

Sub Z_GenErMsg()
Dim Src$(), Mdn, ErNy$(), ErMthSet As Aset
'GoSub T1
'GoSub ZZ1
GoSub ZZ2
Exit Sub
T1:

ZZ2:
    Src = SrczMdn("MXls_Lof_EoLof")
    GoSub Tst
    Brw Act
    Return
ZZ1:
    GoSub Set_Src
    Mdn = "XX"
    GoSub Tst
    Brw Act
    Return
Tst:
    Act = ErMsgzSrc(Src, ErMthSet, Mdn)
    Return
Set_Src:
    Const X$ = "'GenErMsg-Src-Beg." & _
    "|'Val_NotNum      Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number" & _
    "|'Val_NBet      Lno#{Lno} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})" & _
    "|'Val_NotInLis    Lno#{Lno} is [{T1$}] line having invalid Val({ErVal$}).  See valid-value-{VdtValNm$}" & _
    "|'Val_FmlFld      Lno#{Lno} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErFny$()}).  Valid-Fny are [{VdtFny$()}]" & _
    "|'Val_FmlNotBegEq Lno#{Lno} is [Fml] line having [{Fml$}] which is not started with [=]" & _
    "|'Fld_NotInFny    Lno#{Lno} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]" & _
    "|'Fld_Dup         Lno#{Lno} is [{T1$}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno}" & _
    "|'Fldss_NotSel    Lno#{Lno} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]" & _
    "|'Fldss_DupSel    Lno#{Lno} is [{T1$}] line having" & _
    "|'Lon            Lno#{Lno} is [Lo-Nm] line having value({Val$}) which is not a good name" & _
    "|'Lon_Mis        [Lo-Nm] line is missing" & _
    "|'Lon_Dup        Lno#{Lno} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno}" & _
    "|'Tot_DupSel      Lno#{Lno} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored." & _
    "|'Bet_N3Fld        Lno#{Lno} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]" & _
    "|'Bet_EqFmTo      Lno#{Lno} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal." & _
    "|'Bet_FldSeq      Lno#{Lno} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]" & _
    "|'GenErMsg-Src-End." & _
    "|Const M_Bet_FldSeq$ = 1"
    Src = SplitVBar(X)
    Return
End Sub

Function ContFei(Src$(), Ix&) As Fei
ContFei.FmIx = Ix
ContFei.EIx = ContEIx(Src, Ix)
End Function

Function ErMsgzSrc(Src$(), ErMthn As Aset, Optional Mdn = "?") As String()
'Init Src
If ErMthn.IsEmp Then Inf CSub, "No GenErMsg-Src-Beg. / GenErMsg-Src-End.", "Md", Mdn: Exit Function
Dim O$(), O1$(), O2$()
'O1 = SrcRplConstDic(Src, ErConstDic(ErMthn)): 'Brw O1: Stop
O2 = RmvMthInSrc(O1, ErMthn):       'Brw LyzNNAp("MthToRmv BefRmvMth AftRmvMth", ErMthnSet, O1, O2): Stop
'O = AddSy(O2, ErMthLinAy):            'Brw O:Stop
ErMsgzSrc = O
End Function

Function SrcRplConstDic(Src$(), ConstDic As Dictionary) As String()
Dim Cnstn, Dcl$(), Bdy$(), Dcl1$(), Dcl2$()
AsgDclAndBdy Src, Dcl, Bdy
Dcl1 = DclRmvCnstnSet(Dcl, KeySet(ConstDic)): 'Brw Dcl1: Stop
'Brw LyzLinesDicItems(ConstDic): Stop
Dcl2 = Sy(Dcl1, LyzLinesDicItems(ConstDic), Bdy): 'Brw Dcl2: Stop
SrcRplConstDic = Dcl2
End Function

Function DclRmvCnstnSet(Dcl$(), CnstnSet As Aset) As String()
Dim L: For Each L In Itr(Dcl)
    If Not CnstnSet.Has(CnstnzL(L)) Then
        PushI DclRmvCnstnSet, L
    End If
Next
End Function

Sub AsgDclAndBdy(Src$(), ODcl$(), OBdy$())
Dim J&, F&, U&
U = UB(Src)
F = FstMthIx(Src)
If F < 0 Then
    Erase OBdy
    ODcl = Src
    Exit Sub
End If
For J = 0 To F - 1
    PushI ODcl, Src(J)
Next
For J = F To U
    PushI OBdy, Src(J)
Next
End Sub

Function ErMthnSet(ErMthNy$()) As Aset
Set ErMthnSet = AsetzAy(ErMthNy)
End Function

Function ErMthNy(ErNy$()) As String()
Dim I
For Each I In Itr(ErNy)
'    PushI ErMthNy, ErMthn(I)
Next
End Function

Function ErConstDic(ErMthn As Aset, ErMsgAy$()) As Dictionary
Const C$ = "Const  M_?$ = ""?"""
Set ErConstDic = New Dictionary
Dim ErNm, J%
For Each ErNm In ErMthn.Itms
    ErConstDic.Add ErCnstn(ErNm), FmtQQ(C, ErNm, ErMsgAy(J))
    J = J + 1
Next
End Function

Function ErCnstn$(ErNm)
ErCnstn = "M_" & ErNm
End Function

Sub Z_ErMthLinAy()
Dim ErNy$(), ErMsgAy$(), ErMthLinAy$()
'GoSub Z
GoSub T1
Exit Sub
Z:
    Brw ErMthLinAy
    Return
T1:
    ErNy = Sy("Val_NotNum")
    ErMsgAy = Sy("Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number")
    Ept = Sy("Function MsgzVal_NotNum(Lno, T1, Val$) As String(): MsgzVal_NotNum = FmtMacro(M_Val_NotNum, Lno, T1, Val): End Function")
    GoTo Tst
Tst:
    Act = ErMthLinAy
    C
    Return
End Sub

Sub Z_Init()
'Init Src(Md("MXls_Lof_EoLof"))
'If Not HasEle(ErNy, "Bet_FldSeq") Then Stop
'Bet_FldSeq
Stop
End Sub

Function ErMthLinAy(ErNy$(), ErMsgAy$()) As String() 'One ErMth is one-MulStmtLin
Dim J%, O$()
For J = 0 To UB(ErNy)
'    PushI O, ErMthLByNm(ErNy(J), MsgAy(J))
Next
ErMthLinAy = FmtMulStmtSrc(O)
End Function

Function ErMthLByNm$(ErNm$, ErMsg$)
Dim CNm$:         CNm = ErCnstn(ErNm)
Dim ErNy$():     ErNy = NyzMacro(ErMsg)
Dim Pm$:           Pm = JnCommaSpc(AwDist(ErNy))
Dim Calling$: Calling = Jn(AmAddPfx(DclNy(ErNy), ", "))
Dim Mthn:     'Mthn = ErMthn(ErNm)
ErMthLByNm = FmtQQ("Function ?(?) As String():? = FmtMacro(??):End Function", _
    Mthn, Pm, Mthn, CNm, Calling)
End Function

