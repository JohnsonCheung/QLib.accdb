Attribute VB_Name = "MxNDclLin"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxNDclLin."
Public Const FFoNDclLin$ = "Mdn NDclLin"

Function FoNDclLin() As String()
FoNDclLin = SyzSS(FFoNDclLin)
End Function

Function DoNDclLinP() As Drs
DoNDclLinP = DoNDclLinzP(CPj)
End Function

Function DoNDclLinzP(P As VBProject) As Drs
DoNDclLinzP = Drs(FoNDclLin, DyoNDclLin(P))
End Function

Function DyoNDclLin(P As VBProject) As Variant()
Dim C As VBComponent
For Each C In P.VBComponents
    PushI DyoNDclLin, Array(C.Name, NDclLin(C.CodeModule))
Next
End Function

Function NDclLinzS%(Src$())
'Assume FstMth cannot have TopRmk
Dim O&: O = FstMthIx(Src)
NDclLinzS = O - NLinNonCdAbovezS(Src, O)
End Function

Function NDclLin%(M As CodeModule) 'Assume FstMth cannot have TopRmk
Dim O&: O = M.CountOfDeclarationLines
NDclLin = O - NLinNonCdAbove(M, O)
End Function

Function NLinNonCdAbove&(M As CodeModule, Lno&)
Dim O%
Dim J&: For J = Lno To 1 Step -1
    If Not IsLinNCd(M.Lines(J, 1)) Then NLinNonCdAbove = O: Exit Function
    O = O + 1
Next
NLinNonCdAbove = Lno
End Function
Function NLinNonCdAbovezS&(Src$(), Ix&)
Dim O%
Dim J&: For J = Ix To 0 Step -1
    If Not IsLinNCd(Src(J)) Then NLinNonCdAbovezS = O: Exit Function
    O = O + 1
Next
NLinNonCdAbovezS = Ix
End Function

