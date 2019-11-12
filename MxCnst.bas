Attribute VB_Name = "MxCnst"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCnst."
'Enum XoCnst
'    XiMdn
'    XiIsPrv
'    XiCnstn
'    XiTyChr
'    XiCnstv
'End Enum

Function CnstLyP() As String()
CnstLyP = CnstLyzP(CPj)
End Function

Function CnstLyzP(P As VBProject) As String()
CnstLyzP = CnstLy(SrczP(P))
End Function

Function CnstnzL$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfPfxSpc(L, "Const") Then
    CnstnzL = TakNm(L)
End If
End Function

Function CnstLno%(M As CodeModule, Cnstn$, Optional IsPrvOnly As Boolean)
CnstLno = CnstIx(Src(M), Cnstn, IsPrvOnly) + 1
End Function

Function CnstLLinzDcl(Dcl$(), Cnstn$) As LLin
Dim J&: For J = 0 To UB(Dcl)
    Dim L$: L = Dcl(J)
    If CnstnzL(L) = Cnstn Then
        L = ContLin(Dcl, J)
        CnstLLinzDcl = LLin(J + 1, L)
        Exit Function
    End If
Next
End Function

Function CnstLLin(M As CodeModule, Cnstn$) As LLin
CnstLLin = CnstLLinzDcl(DclzM(M), Cnstn)
End Function

Sub Z_HasCnstn()
Debug.Assert HasCnstn(CMd, "CMod")
End Sub

Function HasCnstn(M As CodeModule, Cnstn$) As Boolean
HasCnstn = CnstLno(M, Cnstn) = 0
End Function

Function HasCnstnzL(Lin, N$) As Boolean
HasCnstnzL = CnstnzL(Lin) = N
End Function

Function ShfTermCnst(OLin$) As Boolean
ShfTermCnst = ShfTerm(OLin, "Const")
End Function

Function ShfCnst(OLin$) As Boolean
ShfCnst = ShfT1(OLin) = "Const"
End Function


Function IsLinCnstPfx(L, CnstnPfx$) As Boolean
Dim Lin$: Lin = RmvMdy(L)
If Not ShfTermCnst(Lin) Then Exit Function
IsLinCnstPfx = HasPfx(L, CnstnPfx)
End Function

Private Sub Z_IsLinCnstStr()
Dim O$()
Dim L: For Each L In SrczP(CPj)
    If IsLinCnstStr(L) Then PushI O, L
Next
Brw O
End Sub

Function IsLinCnstStr(Lin) As Boolean
Dim L$: L = Lin
ShfMdy L
If Not ShfTerm(L, "Const") Then Exit Function
If ShfNm(L) = "" Then Exit Function
IsLinCnstStr = FstChr(L) = "$"
End Function

Function IsLinCnst(Lin) As Boolean
Dim L$: L = Lin
ShfMdy L
If Not ShfTerm(L, "Const") Then Exit Function
If ShfNm(L) = "" Then Exit Function
IsLinCnst = True
End Function

Function IsLinCnstzN(L, Cnstn$) As Boolean
IsLinCnstzN = CnstnzL(L) = Cnstn
End Function

Function CnstIx&(Src$(), Cnstn, Optional IsPrvOnly As Boolean)
Dim L, O&
For Each L In Itr(Src)
    If CnstnzL(L) = Cnstn Then
        Select Case True
        Case IsPrvOnly And HasPfx(L, "Public "): CnstIx = -1
        Case Else:                              CnstIx = O
        End Select
        Exit Function
    End If
    O = O + 1
Next
CnstIx = -1
End Function

Function CnstLinAy(Src$()) As String()
Dim Ix&, L: For Each L In Itr(Src)
    If IsLinCnstStr(L) Then PushI CnstLinAy, ContLin(Src, Ix)
    Ix = Ix + 1
Next
End Function

Function CnstLinAyP() As String()
CnstLinAyP = CnstLinAy(SrczP(CPj))
End Function

Sub Z_CnstLy()
Brw CnstLy(SrczP(CPj))
End Sub

Function CnstLy(Src$()) As String()
Dim Ix&: For Ix = 0 To UB(Src)
    If IsLinCnst(Src(Ix)) Then PushI CnstLy, ContLin(Src, Ix)
Next
End Function

