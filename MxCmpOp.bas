Attribute VB_Name = "MxCmpOp"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCmpOp."

Sub AddClsnn(Clsnn$) 'To CPj
AddCmpnnzP Clsnn, vbext_ct_ClassModule, CPj
JmpCmpn T1(Clsnn)
End Sub

Sub AddCmpSfxzP(Sfx, P As VBProject)
If P.Protection = vbext_pp_locked Then Exit Sub
Dim C As VBComponent
For Each C In P.VBComponents
    RenCmp C, C.Name & Sfx
Next
End Sub

Sub AddCmpSfxP(Sfx)
AddCmpSfxzP Sfx, CPj
End Sub

Sub AddCmpzEmp(P As VBProject, Ty As vbext_ComponentType, Nm)
If HasCmpzP(P, Nm) Then InfLin CSub, "Cmpn exist in Pj", "Cmpn Pjn", Nm, P.Name: Exit Sub
P.VBComponents.Add(Ty).Name = Nm ' no CStr will break
End Sub

Sub AddCmpnnzP(Cmpnn$, T As vbext_ComponentType, P As VBProject)
Dim N: For Each N In ItrzSS(Cmpnn)
    AddCmpzEmp P, T, N
Next
End Sub

Sub AddCmpzL(P As VBProject, Cmpn, Srcl$)
AddCmpzEmp P, vbext_ct_StdModule, Cmpn
ApdLines MdzP(P, Cmpn), Srcl
End Sub

Sub AddMod(Modnn$)
AddCmpnnzP Modnn, vbext_ct_StdModule, CPj
JmpCmpn T1(Modnn)
End Sub

Sub AddModnzP(P As VBProject, Modn)
AddCmpzEmp P, vbext_ct_StdModule, Modn
End Sub

Sub ApdLines(M As CodeModule, Lines$)
If Lines = "" Then Exit Sub
M.InsertLines M.CountOfLines + 1, Lines '<=====
End Sub




Sub ApdLineszoInf(M As CodeModule, Lines$)
Dim Bef&, Aft&, Exp&, Cnt&
Bef = M.CountOfLines
ApdLines M, Lines
Aft = M.CountOfLines
Cnt = LinCnt(Lines)
Exp = Bef + Cnt
If Exp <> Aft Then
    Thw CSub, "After copy line count are inconsistents, where [Md], [LinCnt-Bef-Cpy], [LinCnt-of-lines], [Exp-LinCnt-Aft-Cpy], [Act-LinCnt-Aft-Cpy], [Lines]", _
        Mdn(M), Bef, Cnt, Exp, Aft, Lines
End If
End Sub

Sub ApdLy(M As CodeModule, Ly$())
ApdLines M, JnCrLf(Ly)
End Sub

Sub ClrTmpMod()
Dim N
For Each N In TmpModNyzP(CPj)
    If HasPfx(Md(N), "TmpMod") Then RmvCmpzN N
Next
End Sub

Function DftMd(M As CodeModule) As CodeModule
If IsNothing(M) Then
   Set DftMd = CMd
Else
   Set DftMd = M
End If
End Function

Sub DltCmpzPjn(P As VBProject, Mdn)
If Not HasCmpzP(P, Mdn) Then Exit Sub
P.VBComponents.Remove P.VBComponents(Mdn)
End Sub

Sub EnsCls(P As VBProject, Clsn)
EnsCmpzPTN P, vbext_ct_ClassModule, Clsn
End Sub

Sub EnsClsLines(Clsn$, ClsLines$)
EnsCls CPj, Clsn
EnsModLines Md(Clsn), ClsLines
End Sub

Sub EnsCmpzPTN(P As VBProject, Ty As vbext_ComponentType, Nm)
If Not HasCmpzP(P, Nm) Then AddCmpzEmp P, Ty, Nm
End Sub

Sub EnsLines(Md As CodeModule, Mthn, Mthl$)
Dim OldMthL$: OldMthL = MthlzM(Md, Mthn)
If OldMthL = Mthl Then
    Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is same", Mthn, Mdn(Md))
End If
RmvMth Md, Mthn
ApdLines Md, Mthl
Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is replaced <=========", Mthn, Mdn(Md))
End Sub

Sub EnsMod(P As VBProject, Modn)
EnsCmpzPTN P, vbext_ct_StdModule, Modn
End Sub

Sub EnsModLines(M As CodeModule, Lines$)
If Lines = Srcl(M) Then Inf CSub, "Same module lines, no need to replace", "Mdn", Mdn(M): Exit Sub
RplMd M, Lines
End Sub

Sub EnsModzPN(P As VBProject, Mdn)
EnsCmpzPTN P, vbext_ct_StdModule, Mdn
End Sub

Function HasCmpzN(Cmpn) As Boolean
HasCmpzN = HasCmpzP(CPj, Cmpn)
End Function

Function InsDcl(M As CodeModule, Dcl$) As CodeModule
M.InsertLines FstMthLnozM(M), Dcl
Debug.Print FmtQQ("MdInsDcl: Module(?) a DclLin is inserted", Mdn(M))
End Function

Sub RenCmp(A As VBComponent, NewNm$)
If HasCmpzN(NewNm) Then
    InfLin CSub, "New cmp exists", "OldCmp NewCmp", A.Name, NewNm
Else
    A.Name = NewNm
End If
End Sub

Sub RenCmpOfAddPfx(A As VBComponent, AddPfx$)
A.Name = AddPfx & A.Name
End Sub

Sub RenCmpOfRplPfx(A As VBComponent, FmPfx$, ToPfx$)
If HasPfx(A.Name, FmPfx) Then
    A.Name = RplPfx(A.Name, FmPfx, ToPfx)
End If
End Sub

Sub RmvCmp(A As VBComponent)
A.Collection.Remove A
End Sub

Sub RmvCmpzN(Cmpn)
RmvCmp Cmp(Cmpn)
End Sub

Sub RmvMd(MdDn)
RmvMdzMd Md(MdDn)
End Sub

Sub RmvMdzMd(M As CodeModule)
Dim N$, P$
    N = Mdn(M)
    P = PjnzM(M)
'Debug.Print FmtQQ("RmvMd: Before Md(?) is deleted from Pj(?)", M, P)
M.Parent.Collection.Remove M.Parent
Debug.Print FmtQQ("RmvMd: Md(?) is deleted from Pj(?)", N, P)
End Sub

Sub RmvMdzPfx(Pfx$)
Dim Ny$(): Ny = AwPfx(MdNyP, Pfx)
If Si(Ny) = 0 Then InfLin CSub, "no module begins with " & Pfx: Exit Sub
Brw Ny
Dim N
If Cfm("Rmv those Md as show in the notepad?") Then
    For Each N In Ny
        RmvMd Md(N)
    Next
End If
End Sub

Sub RmvModPfx(Pj As VBProject, Pfx$)
Dim C As VBComponent
For Each C In Pj.VBComponents
    If HasPfx(C.Name, Pfx) Then
        RenCmp C, RmvPfx(C.Name, Pfx)
    End If
Next
End Sub

Function SampDiMdnqSrcL() As Dictionary
Set SampDiMdnqSrcL = New Dictionary
Dim C As VBComponent: For Each C In CPj.VBComponents
    SampDiMdnqSrcL.Add C.Name, Srcl(C.CodeModule) & vbCrLf & "'"
Next
End Function

Sub Z_RplMdzD()
RplMdzD CPj, SampDiMdnqSrcL
End Sub


Sub RplMdzD(P As VBProject, DiMdnqSrcL As Dictionary)
'Ret : #Rpl-Md-By-Di-Mdn-SrcL# @@
Dim Mdn: For Each Mdn In DiMdnqSrcL.Keys
    RplMd P.VBComponents(Mdn).CodeModule, DiMdnqSrcL(Mdn)
Next
End Sub

Sub Z_RplMd()
Dim M As CodeModule: Set M = Md("QDao_Def_NewTd")
RplMd M, Srcl(M) & vbCrLf & "'"
End Sub
Function RTrimMd(M As CodeModule, Optional Upd As EmUpd, Optional Osy) As Boolean

End Function

Sub RplMd__ShwMsg(IsSam As Boolean, OldL$, NewL$, Mdn$)
Dim Msg$
    Dim OldC As String * 4: RSet OldC = LinCnt(OldL)
    Msg = Replace("RplMd: OldCnt(?) ", "?", OldC)
    If IsSam Then
        Msg = Msg & "             " & Mdn & vbTab & "<--- Same"
    Else
        Dim NewC As String * 4: RSet NewC = LinCnt(NewL)
        Msg = Msg & Replace("NewCnt(?) ", "?", NewC) & Mdn
    End If
    Debug.Print Msg
End Sub

Function RplMd(M As CodeModule, NewL$) As Boolean
Dim Mdn$: Mdn = M.Name
Dim OldL$: OldL = Srcl(M)
Dim IsSam As Boolean: IsSam = LineszRTrim(OldL) = LineszRTrim(NewL)
RplMd__ShwMsg IsSam, OldL, NewL, Mdn
If IsSam Then Exit Function

ClrMd M:
M.InsertLines 1, NewL
RplMd = True
End Function

Sub RenModPfx(FmPfx$, ToPfx$)
RenModPfxzP CPj, FmPfx, ToPfx
End Sub

Sub RenModPfxzP(Pj As VBProject, FmPfx$, ToPfx$)
Dim C As VBComponent, N$
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        If HasPfx(C.Name, FmPfx) Then
            RenCmp C, RplPfx(C.Name, FmPfx, ToPfx)
        End If
    End If
Next
End Sub

Function SetCmpNm(A As VBComponent, Nm, Optional Fun$ = "SetCmpNm") As VBComponent
Dim Pj As VBProject
Set Pj = PjzC(A)
If HasCmpzP(Pj, Nm) Then
    Thw Fun, "Cmp already Has", "Cmp Has-in-Pj", Nm, Pj.Name
End If
If Pj.Name = Nm Then
    Thw Fun, "Cmpn same as Pjn", "Cmpn", Nm
End If
A.Name = Nm
Set SetCmpNm = A
End Function


Sub ChgToCls(FmModn$)
If Not HasCmp(FmModn) Then InfLin CSub, "Mod not exist", "Mod", FmModn: Exit Sub
If Not IsMod(Md(FmModn)) Then InfLin CSub, "It not Mod", "Mod", FmModn: Exit Sub
Dim T$: T = Left(FmModn & "_" & Format(Now, "HHMMDD"), 31)
Md(FmModn).Name = T
AddClsnn FmModn
Md(FmModn).AddFromString Srcl(Md(T))
RmvCmpzN T
End Sub
