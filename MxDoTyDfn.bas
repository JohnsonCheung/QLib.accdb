Attribute VB_Name = "MxDoTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoTyDfn."

Function DoTyDfnP() As Drs
':DoTyDfn: :Drs-Mdn-Nm-Ty-Mem-Rmk
DoTyDfnP = DoTyDfnzP(CPj)
End Function

Function DoTyDfnzP(P As VBProject) As Drs
Dim O As Drs
Dim C As VBComponent: For Each C In P.VBComponents
    O = AddDrs(O, DoTyDfnzCmp(C))
Next
DoTyDfnzP = O
End Function

Function DoTyDfnzCmp(C As VBComponent) As Drs
Dim S$(): S = Src(C.CodeModule)
Dim Dy(): Dy = DyoTyDfn(VbRmk(S), C.Name)
DoTyDfnzCmp = Drs(FoTyDfn, Dy)
End Function

Function FoTyDfn() As String()
FoTyDfn = SyzSS(FFoTyDfn)
End Function

Function DyoTyDfnzP(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy DyoTyDfnzP, DyoTyDfnzM(C.CodeModule)
Next
End Function

Function DyoTyDfnzM(M As CodeModule) As Variant()
DyoTyDfnzM = DyoTyDfn(VbRmk(Src(M)), Mdn(M))
End Function

Function DyoTyDfn(VbRmk$(), Mdn$) As Variant()
':DyoTyDfn: :Dyo-Nm-Ty-Mem-VbRmk ! Fst-Lin must be :nn: :dd #mm# !rr
'                                ! Rst-Lin is !rr
'                                ! must term: nn dd mm, all of them has no spc
'                                ! opt      : mm rr
'                                ! :xx:     : should uniq in pj
Dim G(): G = TyDfnGp(VbRmk)
Dim Gp: For Each Gp In Itr(G)
    PushI DyoTyDfn, DroTyDfn(CvSy(Gp), Mdn)
Next
End Function

Function TyDfnGp(VbRmk$()) As Variant()
Dim O()
Dim L: For Each L In Itr(VbRmk)
    Dim NFstLin%
    Dim Gp()
    Select Case True
    Case IsLinTyDfn(L)
        NFstLin = NFstLin + 1
        PushSomSi O, Gp
        Erase Gp
        PushI Gp, L
    Case IsLinTyDfnRmk(L)
        If Si(Gp) > 0 Then
            PushI Gp, L ' Only with Fst-Lin, the Rst-Lin will be use, otherwise ign it.
        End If
    Case Else
        PushSomSi O, Gp
        Erase Gp
    End Select
Next
TyDfnGp = O
End Function

Sub Z_DroTyDfn()
Dim VbRmk$()
GoSub ZZ
Exit Sub
ZZ:
    VbRmk = Sy("':Cell: :SCell-or-:WCell")
    Dmp DroTyDfn(VbRmk, "Md")
    Return
End Sub

Function DroTyDfn(TyDfnLy$(), Mdn$) As Variant()
'Assume: Fst Lin is ':nn: :dd [#mm#] [!rr]
'        Rst Lin is '                 !rr
Dim Dr(): Dr = DroTyDfnzL(TyDfnLy(0), Mdn)
Dr(4) = AddNB(Dr(4), ExmRmkl(CvSy(RmvFstEle(TyDfnLy))))
DroTyDfn = Dr
End Function

Function DroTyDfnzL(FstTyDfnLin$, Mdn$) As Variant()
Dim L$: L = FstTyDfnLin
Dim Nm$, Ty$, Mem$, Rmk$
Nm = ShfTyDfnNm(L)
If Nm = "" Then Exit Function
Nm = RmvFstChr(Nm)
Ty = ShfDfnTy(L)
Mem = ShfMemNm(L)
If L <> "" Then
    If FstChr(L) <> "!" Then Thw CSub, "Given FstTyDfnLin is not in valid format", "FstTyDfnLin", FstTyDfnLin
    Rmk = Trim(RmvFstChr(L))
End If
DroTyDfnzL = Array(Mdn, Nm, Ty, Mem, Rmk)
End Function
