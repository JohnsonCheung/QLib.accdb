Attribute VB_Name = "MxIsLin"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIsLin."
Function IsLinPrp(L) As Boolean
IsLinPrp = MthKd(L) = "Property"
End Function

Sub Z_IsLinMth()
GoTo Z
Dim A$
A = "Function IsLinMth(A) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = IsLinMth(A)
    C
    Return
Z:
Dim L, O$()
For Each L In CSrc
    If IsLinMth(CStr(L)) Then
        PushI O, L
    End If
Next
Brw O
End Sub

Function HitShtMdy(ShtMdy$, ShtMthMdyAy$()) As Boolean
HitShtMdy = HitAy(IIf(ShtMdy = "", "Pub", ShtMdy), ShtMthMdyAy)
End Function

Function IsLinOptOrImplOrBlnk(L) As Boolean
IsLinOptOrImplOrBlnk = True
If IsLinOpt(L) Then Exit Function
If IsLinImpl(L) Then Exit Function
If L = "" Then Exit Function
IsLinOptOrImplOrBlnk = False
End Function

Function IsLinImpl(L) As Boolean
IsLinImpl = HasPfx(L, "Implements ")
End Function

Function IsLinOpt(L) As Boolean
If Not HasPfx(L, "Option ") Then Exit Function
Select Case True
Case _
    HasPfx(L, "Option Explicit"), _
    HasPfx(L, "Option Compare Text"), _
    HasPfx(L, "Option Compare Binary"), _
    HasPfx(L, "Option Compare Database")
    IsLinOpt = True
End Select
End Function

Function IsLinPubMth(L) As Boolean
Dim Lin$: Lin = L
Dim Mdy$: Mdy = ShfMdy(Lin): If Mdy <> "" And Mdy <> "Public" Then Exit Function
IsLinPubMth = TakMthKd(Lin) <> ""
End Function

Function IsMthLinSngL(L) As Boolean
Dim K$: K = MthKd(L): If K = "" Then Exit Function
IsMthLinSngL = HasSubStr(L, "End " & K)
End Function

Function IsLinMth(L) As Boolean
IsLinMth = MthKd(L) <> ""
End Function
Function IsLinMthNm(L, Nm) As Boolean
IsLinMthNm = Mthn(L) = Nm
End Function

Function IsLinEmn(A) As Boolean
IsLinEmn = HasPfx(RmvMdy(A), "Enum ")
End Function

Function IsLinTy(A) As Boolean
IsLinTy = HasPfx(RmvMdy(A), "Type ")
End Function

Function IsLinEmpSrc(A) As Boolean
IsLinEmpSrc = True
If HasPfx(A, "Option ") Then Exit Function
Dim L$: L = Trim(A)
If L = "" Then Exit Function
IsLinEmpSrc = False
End Function

Function IsLinSngTerm(L) As Boolean
IsLinSngTerm = InStr(Trim(L), " ") = 0
End Function

Function IsLinDD(L) As Boolean
IsLinDD = Fst2Chr(LTrim(L)) = "--"
End Function
