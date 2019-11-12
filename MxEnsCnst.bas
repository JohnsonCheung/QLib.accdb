Attribute VB_Name = "MxEnsCnst"
Option Explicit
Option Compare Text
Const CLib$ = "QIde"
Const CMod$ = CLib & "MxEnsCnst."

Sub EnsCnstLin(M As CodeModule, CnstLin$)
Dim Lno&: Lno = CnstLno(M, CnstnzL(CnstLin))
If Lno > 0 Then
    If M.Lines(Lno, 1) = CnstLin Then Exit Sub
    M.ReplaceLine Lno, CnstLin
    Exit Sub
End If
InsCnstLin M, CnstLin
End Sub

Sub EnsCnstLinAft(M As CodeModule, CnstLin$, AftCnstn$, Optional IsPrvOnly As Boolean)
Dim Lno&: Lno = CnstLno(M, CnstnzL(CnstLin))
If IsPrvOnly Then
    If Lno > 0 Then
        If HasPfx(M.Lines(Lno, 1), "Public ") Then
            Exit Sub
        End If
    End If
End If
If Lno > 0 Then
    If M.Lines(Lno, 1) = CnstLin Then Exit Sub
    M.ReplaceLine Lno, CnstLin
    InfLin CSub, "CnstLin is replaced", "Mdn CnstLin", Mdn(M), CnstLin
    Exit Sub
End If
InsCnstLinAft M, CnstLin, AftCnstn
End Sub

Sub ClrMdLin(M As CodeModule, Lno&, Optional Cnt& = 1)
Dim J&: For J = Lno To Lno + Cnt - 1
    M.ReplaceLine Lno, ""
Next
End Sub

Sub InsCnstLin(M As CodeModule, CnstLin$)
Dim Lno&: Lno = LnoAftOptqImpl(M)
M.InsertLines Lno, CnstLin
InfLin CSub, "CnstLin is inserted", "Lno Mdn CnstLin", Lno, Mdn(M), CnstLin
End Sub

Sub InsCnstLinAft(M As CodeModule, CnstLin$, AftCnstn$)
Dim Lno&
    Lno = CnstLno(M, AftCnstn): If Lno <> 0 Then Lno = Lno + 1
    If Lno = 0 Then Lno = LnoAftOptqImpl(M)
M.InsertLines Lno, CnstLin
InfLin CSub, "CnstLin is inserted", "Lno Mdn CnstLin", Lno, Mdn(M), CnstLin
End Sub

Sub ClrCnstLin(M As CodeModule, Cnstn$)
Dim Lno&: Lno = CnstLno(M, Cnstn)
If Lno > 0 Then
    M.ReplaceLine Lno, ""
    InfLin CSub, "Cnstn is cleared", "Mdn Cnstn", Mdn(M), Cnstn
End If
End Sub

Sub RmvCnstLin(M As CodeModule, Cnstn$, Optional IsPrvOnly As Boolean)
Dim Lno&: Lno = CnstLno(M, Cnstn, IsPrvOnly)
If Lno > 0 Then
    M.DeleteLines Lno, 1
    InfLin CSub, "Cnstn is removed", "Mdn Cnstn", Mdn(M), Cnstn
End If
End Sub

Sub RmvCnstLinzP(P As VBProject, Cnstn$, Optional IsPrvOnly As Boolean)
Dim C As VBComponent: For Each C In P.VBComponents
    RmvCnstLin C.CodeModule, Cnstn, IsPrvOnly
Next
End Sub
