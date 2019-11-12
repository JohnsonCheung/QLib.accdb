Attribute VB_Name = "MxAlignLy"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAlignLy."
Enum EmAlign
    EiLeft
    EiRight
End Enum

Function AlignLyz1T(L$()) As String()
AlignLyz1T = AlignLyzNTerm(L, 1)
End Function

Function AlignLyz2T(L$()) As String()
AlignLyz2T = AlignLyzNTerm(L, 2)
End Function

Function AlignLyz3T(L$()) As String()
AlignLyz3T = AlignLyzNTerm(L, 3)
End Function

Function AlignLyz4T(L$()) As String()
AlignLyz4T = AlignLyzNTerm(L, 4)
End Function

Function AlignLyzNTerm(L$(), N%) As String()
Dim Dy(): Dy = LyMapNTermRst(L, N)
Dim Dy1(): Dy1 = AlignDy(Dy, N)
AlignLyzNTerm = AmRTrim(JnDy(Dy1))
End Function

Function WdtAyzFstNTerm(NTerm%, L$()) As Integer()
If Si(L) = 0 Then Exit Function
Dim O%(), W%(), I
ReDim O(NTerm - 1)
For Each I In Itr(L)
    W = WdtAyzFstNTermL(NTerm, L)
    O = WdtAyz2W(O, W)
Next
WdtAyzFstNTerm = O
End Function

Function LyMapSplitDot(L$()) As Variant()
Dim I: For Each I In Itr(L)
    PushI LyMapSplitDot, SplitDot(I)
Next
End Function

Function LyMapNTermRst(L$(), N%) As Variant()
Dim I: For Each I In Itr(L)
    PushI LyMapNTermRst, NTermRst(I, N)
Next
End Function

Function WdtAyzFstNTermL(N%, Lin) As Integer()
Dim T
For Each T In FstNTerm(Lin, N)
    PushI WdtAyzFstNTermL, Len(T)
Next
End Function

Function WdtAyz2W(W1%(), W2%()) As Integer()
Dim O%(): O = W1
Dim I, J%: For Each I In W2
    If I > O(J) Then O(J) = I
    J = J + 1
Next
WdtAyz2W = O
End Function

Function AlignLyzDot(Ly_wi_Dot$()) As String()
AlignLyzDot = FmtDy(LyMapSplitDot(Ly_wi_Dot))
End Function

Sub BrwDotLy(DotLy$())
Brw AlignDotLy(DotLy)
End Sub

Function AlignDotLy(DotLy$()) As String()
AlignDotLy = FmtDy(DyoDotLy(DotLy), Fmt:=EiSSFmt)
End Function

Function AlignDotLyzTwoCol(DotLy$()) As String()
AlignDotLyzTwoCol = FmtDy(DyoDotLyzTwoCol(DotLy), Fmt:=EiSSFmt)
End Function

Sub Z_AlignAy2T()
Dim L$()
L = Sy("AAA B C D", "L BBB CCC")
Ept = Sy("AAA B   C D", _
         "L   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = AlignLyz2T(L)
    C
    Return
End Sub
Sub Z_AlignLyz3T()
Dim L$()
L = Sy("AAA B C D", "L BBB CCC")
Ept = Sy("AAA B   C   D", _
         "L   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = AlignLyz3T(L)
    C
    Return
End Sub

