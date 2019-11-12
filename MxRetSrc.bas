Attribute VB_Name = "MxRetSrc"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxRetSrc."

Function SrclFi$(M As CodeModule, b As Fei)
If IsFeizEmp(b) Then Exit Function
'SrclFi = A.Lines(B.FmNo, B.Cnt)
End Function

Function SrczMFi(M As CodeModule, b As Fei) As String()
'SrczMFi = SplitCrLf(SrclFei(A, B))
End Function

Function SrczMRe(M As CodeModule, b As RegExp) As String()
Dim Ix&(): Ix = IxyzAyRe(Src(M), b)
Dim O$(), I, Md As CodeModule
Dim N$: N = Mdn(M)
If Si(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdJmpLin ""?"",??' ?", N, I + 1, vbTab, M.Lines(I + 1, 1))
Next
SrczMRe = O
End Function

Function SrczPjPatn(P As VBProject, Patn$) As String()
SrczPjPatn = AwPatn(SrczP(P), Patn)
End Function
