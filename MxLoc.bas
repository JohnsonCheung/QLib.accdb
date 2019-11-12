Attribute VB_Name = "MxLoc"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxLoc."

Sub JmpLin_InMd(M As CodeModule, Lno&)
JmpMd M
JmpLno Lno
End Sub

Function Has_Mdn_InPj(P As VBProject, Mdn, Optional Inf As Boolean) As Boolean
If HasItn(CPj.VBComponents, Mdn) Then Has_Mdn_InPj = True: Exit Function
If Inf Then InfLin CSub, FmtQQ("Md[?] not exist", Mdn)
End Function

Function Has_Mdn_InCurPj(Mdn, Optional Inf As Boolean) As Boolean
Has_Mdn_InCurPj = Has_Mdn_InPj(CPj, Mdn, Inf)
End Function

Sub Z_Jmp()
Jmp "QIde_Md.10"
End Sub

Sub Jmp(MdnDotLno$)
JmpMdn BefOrAll(MdnDotLno, ".")
JmpLin Aft(MdnDotLno, ".")
End Sub

Sub JmpRCC_InCurMd(R&, C1%, C2%)
CPne.SetSelection R, C1, R, C2
End Sub

Sub JmpMdn(Mdn)
If Not Has_Mdn_InCurPj(Mdn) Then Debug.Print "Mdn not exist": Exit Sub
JmpMd Md(Mdn)
End Sub

Sub JmpLno(Lno&)
Dim C2%: C2 = Len(CMd.Lines(Lno, 1)) + 1
JmpLcc Lno, 1, C2
End Sub

Sub JmpLcc(Lno&, C1%, C2%)
With CPne
    .TopLine = Lno
    .SetSelection Lno, C1, Lno, C2
End With
End Sub

Sub JmpLin(MdLnoOrMdLnoCCStr$)
':MdLnoStr: :Term ! Mdn:Lno
Dim Mdn$, LnoOrLnoCCStr$
AsgBrk MdLnoOrMdLnoCCStr, ":", Mdn, LnoOrLnoCCStr
JmpMdn Mdn
If HasSubStr(LnoOrLnoCCStr, ":") Then
    Dim A$(): A = SplitColon(LnoOrLnoCCStr)
    Dim Lno&, C1%, C2%
    Lno = A(0)
    C1 = A(1)
    C2 = A(2)
    JmpLcc Lno, C1, C2
Else
    JmpLno CLng(LnoOrLnoCCStr)
End If
End Sub

Sub JmpRRCC(A As RRCC)
Dim L&, C1%, C2%
With CPne
    If C1 = 0 Or C2 = 0 Then
        C1 = 1
        C2 = Len(.CodeModule.Lines(L, 1)) + 1
    End If
    .TopLine = L
    .SetSelection L, C1, L, C2
End With
'SendKeys "^{F4}"
End Sub

Sub JmpMthzMN(M As CodeModule, Mthn)
JmpMd M
JmpMth Mthn
End Sub

Sub JmpPj(P As VBProject)
ClsAllWin
Dim M As CodeModule
Set M = FstMd(P)
If IsNothing(M) Then Exit Sub
JmpMd M
TileV
DoEvents
End Sub

Sub JmpMdRRCC(M As CodeModule, R As RRCC)
JmpMd M
JmpRRCC R
End Sub

Function WinyzMdAy(MdAy() As CodeModule) As VBIDE.Window()

End Function

Sub JmpMdnn(Mdnn$)
Dim MdAy() As CodeModule: MdAy = MdAyzNN(Mdnn)
Dim I
For Each I In Itr(MdAy)
    JmpMd CvMd(I)
Next
ClsWinExlAp WinyzMdAy(MdAy)
TileV
End Sub

Sub JmpMd(M As CodeModule)
M.CodePane.Show
End Sub
