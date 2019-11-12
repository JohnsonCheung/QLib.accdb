Attribute VB_Name = "MxEnsAsm"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEnsAsm."
':Osy: :Missing|Sy #Oup-String-Array# ! if given as Sy, always push the msg into it.  Always together and aft Upd:EmpUpd.
Function NsNm$(Mdn$)
NsNm = AftRev(Mdn, "_")
End Function

Function EnsNsNmzM(M As CodeModule, Optional Upd As EmUpd, Optional Osy) As Boolean
'Ret : :Msg # Only when Rpt=EiPushOnly or EiUpdAndPush
'       Add/Rpl Line - Const  NsNm$ = "<NsNm>" if NsNm='', no upd, just shw msg.  @@
If IsMdEmp(M) Then Exit Function
If CmpTyzM(M) = vbext_ct_Document Then Exit Function

Dim L As LLin: 'LLin = ConstLLin(M, "NsNm")

Dim OLno%
    If L.Lno = 0 Then
        OLno = 1
        Stop
    Else
        OLno = L.Lno
    End If

Dim Mdn$: Mdn = MdnzM(M)
Const T$ = "Const  ?$ = ""?"""
Dim ONewL$: ONewL = FmtQQ(T, NsNm(Mdn))

Dim OldL$: OldL = L.Lin

Dim OIsIns As Boolean:    OIsIns = OldL = ""

Dim OIsRpl As Boolean:    OIsRpl = OldL <> "" And OldL <> ONewL

'== RplLin =============================================================================================================
'   InsLin
If IsUpd(Upd) Then
    If OIsIns Then M.InsertLines OLno, ONewL
    If OIsRpl Then M.ReplaceLine OLno, ONewL
End If

    Dim IsRpt As Boolean, IsPush As Boolean
    IsRpt = IsRptU(Upd)
    IsPush = IsRptU(Upd)
    If IsRpt Or IsPush Then
        Dim Msg$: Msg = MsgzEnsNs(OIsIns, OIsRpl, OLno, ONewL)
        If IsRpt Then Brw Msg
        If IsPush Then EnsNsNmzM = Msg
    End If
End Function

Function MsgzEnsNs$(IsIns As Boolean, IsRpl As Boolean, Lno%, NewL$)
Stop
End Function

Sub EnsNsNmzP(P As VBProject, Optional Upd As EmUpd, Optional Osy)
Dim C As VBComponent, Mdyd%, Skpd%
For Each C In P.VBComponents
    If EnsNsNmzM(C.CodeModule, Upd, Osy) Then
        Mdyd = Mdyd + 1
    Else
        Skpd = Skpd + 1
    End If
'    Brw XX: Stop
Next
Brw XX
Inf CSub, "Done", "Pj Mdyd Skpd Tot", P.Name, Mdyd, Skpd, Mdyd + Skpd
End Sub

Function XRpl(Lno&, LAct$, LEpt$) As Drs
If Lno = 0 Then Exit Function
If LAct = LEpt Then Exit Function
XRpl = LNewO(Av(Array(Lno, LEpt, LAct)))
End Function

Function LnozDclCnst%(M As CodeModule, Cnstn$)
Dim O%, L$
Dim C$: C = "Const " & Cnstn
For O = 1 To M.CountOfDeclarationLines
    L = RmvMdy(M.Lines(O, 1))
    If ShfPfx(L, "Const ") Then
        If TakNm(L) = Cnstn Then LnozDclCnst = O: Exit Function
    End If
Next
End Function

Function LnozFstCd&(M As CodeModule)
Stop

End Function
Function LnozFstDcl&(M As CodeModule)
Dim J&
For J = 1 To M.CountOfDeclarationLines
    Dim L$: L = Trim(M.Lines(J, 1))
    If Not HasPfxss(L, "Option Implements '") Then
        If L <> "" Then
            LnozFstDcl = True
            Exit Function
        End If
    End If
Next

End Function

Sub Z_LnozDclConst()
Dim Md As CodeModule, Cnstn$
GoSub T0
Exit Sub
T0:
    Set Md = CMd
    Cnstn = "A$"
    Ept = 14&
    GoTo Tst
Tst:
    Act = LnozDclCnst(Md, Cnstn)
    C
    Return
End Sub

Sub Z_NsNm()
BrwDrs DrszMapAy(Itn(CPj.VBComponents), "NsNm", , "Mdn")
End Sub

