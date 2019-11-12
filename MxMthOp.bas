Attribute VB_Name = "MxMthOp"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthOp."

Function RplMth(M As CodeModule, Mthn, NewL$) As Boolean
'Ret : True if Replaced.  Will dlt if NewL=''
Dim Lno&: Lno = MthLnozMM(M, Mthn)
If Not HasMthzM(M, Mthn) Then
    RplMth = True
    If NewL <> "" Then
        M.InsertLines M.CountOfLines + 1, NewL '<=== added at end
    End If
    Exit Function
End If
Dim OldL$: OldL = MthlzM(M, Mthn)
If OldL = NewL Then Exit Function '<== no chg
RplMth = True
RmvMth M, Mthn '<== rmv
M.InsertLines Lno, NewL '<== and ins
End Function

Sub RmvMthzNN(M As CodeModule, Mthnn)
Dim I
For Each I In TermAy(Mthnn)
    RmvMth M, I
Next
End Sub

Sub RmvMth(M As CodeModule, Mthn)
With MthLnoC2(M, Mthn)
    If .S2 > 0 Then M.DeleteLines .S2, .C2
    If .S1 > 0 Then M.DeleteLines .S1, .C1
End With
End Sub

Sub CpyMthAs(M As CodeModule, Mthn, AsMthn)
If HasMthzM(M, AsMthn) Then Inf CSub, "AsMth exist.", "Mdn FmMth AsMth", Mdn(M), Mthn, AsMthn: Exit Sub
End Sub

Sub Z_RmvMth()
Dim Md As CodeModule
Const Mthn$ = "ZZRmv1"
Dim Bef$(), Aft$()
Crt:
        Set Md = TmpMod
        ApdLines Md, LineszVbl("|'sdklfsdf||'dsklfj|Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Sub SetZZRmv1(V)|End Property")
Tst:
        Bef = Src(Md)
        RmvMth Md, Mthn
        Aft = Src(Md)

Insp:   Insp CSub, "RmvMth Test", "Bef RmvMth Aft", Bef, Mthn, Aft
Rmv:    RmvMd Md
End Sub


Sub MovMthzNM(Mthn, ToMdn)
MovMthzMNM CMd, Mthn, Md(ToMdn)
End Sub

Sub MovMthzMNM(Md As CodeModule, Mthn, ToMd As CodeModule)
CpyMth Mthn, Md, ToMd
RmvMth Md, Mthn
End Sub

Function CdzEmpFun$(FunNm)
CdzEmpFun = FmtQQ("Function ?()|End Function", FunNm)
End Function

Function CdzEmpSub$(Subn)
CdzEmpSub = FmtQQ("Sub ?()|End Sub", Subn)
End Function

Sub AddSub(Subn)
ApdLines CMd, CdzEmpSub(Subn)
JmpMth Subn
End Sub

Sub JmpMth(Mthn)
Dim M As CodeModule: Set M = MdzMthn(CPj, Mthn)
JmpMd M
JmpLno MthLnozMM(M, Mthn)
End Sub

Sub AddFun(FunNm)
ApdLines CMd, CdzEmpFun(FunNm)
JmpMth FunNm
End Sub

Function CpyMth(Mthn, FmM As CodeModule, ToM As CodeModule) As Boolean
Dim NewL$
'NewL
    NewL = MthlzM(FmM, Mthn)
'Rpl
    CpyMth = RplMth(ToM, Mthn, NewL)
'
'Const CSub$ = CMod & "CpyMdMthToMd"
'Dim Nav(): ReDim Nav(2)
'GoSub BldNav: ThwIf_ObjNE Md, ToMd, CSub, "Fm & To md cannot be same", Nav
'If Not HasMthzM(Md, Mthn) Then GoSub BldNav: ThwNav CSub, "Fm Mth not Exist", Nav
'If HasMthzM(ToMd, Mthn) Then GoSub BldNav: ThwNav CSub, "To Mth exist", Nav
'ToMd.AddFromString MthlzM(Md, Mthn)
'RmvMth Md, Mthn
'If Not IsSilent Then Inf CSub, FmtQQ("Mth[?] in Md[?] is copied ToMd[?]", Mthn, Mdn(Md), Mdn(ToMd))
'Exit Sub
'BldNav:
'    Nav(0) = "FmMd Mth ToMd"
'    Nav(1) = Mdn(Md)
'    Nav(2) = Mthn
'    Nav(3) = Mdn(ToMd)
'    Return
End Function

Function CpyMthAsVer(M As CodeModule, Mthn, Ver%) As Boolean
'Ret True if copied
Dim VerMthn$, NewL$, L$, OldL$
If Not HasMthzM(M, Mthn) Then Inf CSub, "No from-mthn", "Md Mthn", Mdn(M), Mthn: Exit Function
VerMthn = Mthn & "_Ver" & Ver
'NewL
    L = MthlzM(M, Mthn)
    NewL = Replace(L, "Sub " & Mthn, "Sub " & VerMthn, Count:=1)
'Rpl
    CpyMthAsVer = RplMth(M, VerMthn, NewL)

End Function
