Attribute VB_Name = "MxMdCache"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMdCache."

Function IsCachedM() As Boolean
IsCachedM = IsCachedzM(CMd)
End Function

Function IsCachedzM(M As CodeModule) As Boolean
Const CSub$ = CMod & "IsCachedzM"
Const Trc As Boolean = False
Dim Cs$(): Cs = CachedSrczM(M)
Dim Clc&: Clc = Si(Cs)
Dim Mlc&: Mlc = M.CountOfLines + 1
If Mlc <> Clc Then
    If Trc Then InfLin CSub, "LinCnt dif", "Md-LinCnt Src-LinCnt", Mlc, Clc
    Exit Function
End If
If Clc = 0 Then IsCachedzM = True
Dim Ms$(): Ms = Src(M)
Push Ms, ""
IsCachedzM = IsEqSy(Cs, Ms)
End Function

Function CachedSrcM() As String()
CachedSrcM = CachedSrczM(CMd)
End Function

Function CachedSrczM(M As CodeModule) As String()
Dim F$: F = SrcFfn(M.Parent)
If NoFfn(F) Then Exit Function
Dim S$(): S = LyzFt(F)
Dim S1$(): S1 = RmvClsSig(S)
CachedSrczM = RmvAtrVB(S1)

End Function

Function RmvAtrVB(S$()) As String()
Dim N&: N = AtrVBCnt(S$())
RmvAtrVB = AeFstNEle(S, N)
End Function

Function AtrVBCnt%(S$())
Dim O%:
    Dim L: For Each L In Itr(S)
        If NoPfx(L, "Attribute VB") Then Exit For
        O = O + 1
    Next
AtrVBCnt = O
End Function

Function RmvClsSig(S$()) As String()
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
If HasClsSig(S) Then
    RmvClsSig = AeFstNEle(S, 4)
Else
    RmvClsSig = S
End If
End Function

Function HasClsSig(S$()) As Boolean
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
If Si(S) < 4 Then Exit Function
If S(0) <> "VERSION 1.0 CLASS" Then Exit Function
If S(1) <> "BEGIN" Then Exit Function
If HasPfx(S(2), "  MultiUse =") Then Exit Function
If S(3) = "End" Then Exit Function
End Function
