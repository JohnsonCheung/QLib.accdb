Attribute VB_Name = "MxSqTpEr"
Option Compare Text
Option Explicit
Const CLib$ = "QTp."
Const CMod$ = CLib & "MxSqTpEr."
Function EoSqTp(SqTp$) As String()

End Function

Function EoSqLy(SqLy$()) As LyRslt

End Function
Function MsgAp_Lin_TyEr(DroLLin()) As String()


End Function

Function MsgMustBeIntoLin$(DroLLin())

End Function

Function MsgMustBeSelorSelDis$(DroLLin())

End Function

Function MsgMustNotHasSpcInTbl_NmOfIntoLin$(DroLLin())

End Function
Function BlkIx%(B As Blk)
BlkIx = B.DroBlk(3)
End Function
Function EoExcessBlk(B As Blks, BlkTy$) As String()
Dim M As Blk: 'M = BlkswTy(B, BlkTy)
If IsBlkEmp(M) Then Exit Function
PushI EoExcessBlk, FmtQQ("Excess [?] block, they are ignored", BlkTy)
'PushI EoExcessBlk, EoAftBlk(B, M)
End Function

Function MsgzLeftOvrAftEvl(A() As SwLin, Sw As Sw) As String()
'If Si(A) = 0 Then Exit Function
Dim I
PushI MsgzLeftOvrAftEvl, "Following lines cannot be further evaluated:"
'For Each I In A
'    PushI MsgzLeftOvrAftEvl, vbTab & CvSwLin(I).Lin
'Next
'PushIAy MsgzLeftOvrAftEvl, FmtDicTit(Sw, "Following is the [Sw] after evaluated:")
End Function

Function MsgzNoNm$(A As SwLin)
'MsgzNoNm = SwLinMsg(A, "No name")
End Function

Function MsgzOpStrEr$(A As SwLin)
MsgzOpStrEr = SwLinMsg(A, "2nd Term [Op] is invalid operator.  Valid operation [NE EQ AND OR]")
Stop
End Function

Function MsgzPfx$(A As SwLin)
MsgzPfx = SwLinMsg(A, "First Char must be @")
End Function
Function SwLinMsg$(A As SwLin, ParamArray Ap())

End Function
Function MsgzTermCntAndOr$(A As SwLin)
MsgzTermCntAndOr = SwLinMsg(A, "When 2nd-Term (Operator) is [AND OR], at least 1 term")
End Function

Function MsgzTermCntEqNe$(A As SwLin)
MsgzTermCntEqNe = SwLinMsg(A, "When 2nd-Term (Operator) is [EQ NE], only 2 terms are allowed")
End Function

Function MsgzTermMustBegWithQuestOrAt$(TermAy, A As SwLin)
MsgzTermMustBegWithQuestOrAt = SwLinMsg(A, "Terms[" & JnSpc(TermAy) & "] must begin with either [?] or [@?]")
End Function

Function MsgzTermNotInPm$(TermAy, A As SwLin)
MsgzTermNotInPm = SwLinMsg(A, "Terms[" & JnSpc(TermAy) & "] begin with [@?] must be found in Pm")
End Function

Function MsgzTermNotInSw$(TermAy, A As SwLin, SwNm As Dictionary)
MsgzTermNotInSw = SwLinMsg(A, "Terms[" & JnSpc(TermAy) & "] begin with [?] must be found in Switch")
End Function

Function EoDupNm(A() As SwLin, O() As SwLin) As String()
Dim Ny$(), Nm$
Dim J%, M As SwLin, Er() As SwLin
'For J = 0 To UB(A)
    'Set M = A(J)
    If HasEle(Ny, M.Nm) Then
        'PushObj Er, M
    Else
        'PushObj O, M
        PushI Ny, M.Nm
    End If
'Next
'EoDupNm = MsgzDupNm(Er)
End Function

Function EoFld(A() As SwLin, O() As SwLin) As String()
Exit Function
Dim M As SwLin, IsEr As Boolean, J%, I, A1() As SwLin, A2() As SwLin
IsEr = True
A1 = A
While IsEr
    J = J + 1: If J > 1000 Then Stop
    IsEr = False
'    For Each I In Itr(A1)
'        Set M = EoFldLin(CvSwLin(I), SwNmDic(A), OEr)
'        If IsNothing(M) Then
            IsEr = True
'        Else
            'PushObj A2, M
'        End If
    'Next
    A1 = A2
Wend
'EoFld = A2
End Function

Function EoFldLin(A As SwLin, SwNm As Dictionary, Pm As Dictionary, OEr$()) As SwLin
'Each Term in A.TermAy must be found either in Sw or Pm
Dim O0$(), O1$(), O2$(), I
For Each I In Itr(A.TermAy)
    Select Case True
    Case HasPfx(I, "?"):  If Not SwNm.Exists(I) Then Push O0, I
    Case HasPfx(I, "@?"): If Not Pm.Exists(I) Then Push O1, I
    Case Else:                  Push O2, I
    End Select
Next
PushIAy OEr, MsgzTermNotInSw(O0, A, SwNm)
PushIAy OEr, MsgzTermNotInPm(O1, A)
PushIAy OEr, MsgzTermMustBegWithQuestOrAt(O2, A)
'If HasElezInSomAyzOfAp(O0, O1, O2) Then Set EoFldLin = A
End Function

Function EoLin1$(IO As SwLin)
Dim SwLinMsg$
With IO
    If .Nm = "" Then EoLin1 = MsgzNoNm(IO): Exit Function
'    Select Case .OpStr
'    Case "OR", "AND": If Si(.TermAy) = 0 Then EoLin1 = MsgzTermCntAndOr(IO): Exit Function
'    Case "EQ", "NE":  If Si(.TermAy) <> 2 Then EoLin1 = MsgzTermCntEqNe(IO): Exit Function
'    Case Else:        EoLin1 = MsgzOpStrEr(IO): Exit Function
'    End Select
End With
End Function

Function EoPm(Pm As Dictionary) As String()
'Dim O As Dictionary: Set O = Dic(PmLy)
'Dim B As Boolean, K, V
'For Each K In O.Keys
'    Select Case True
'    Case HasPfx(K, ">?"):
'        V = O(K)
'        Select Case V
'        Case "0": B = False
'        Case "1": B = True
'        Case Else: Thw CSub, "If K='>?xxx', V should be 0 or 1", "K V PmLy", K, V, PmLy
'        End Select
'        O(K) = B
'    Case HasPfx(K, ">")
'    Case Else: Thw CSub, "Pm line should beg with (>? | >)", "K V PmLy", K, O(K), PmLy
'    End Select
'Next
'Set SqPm = O
End Function
