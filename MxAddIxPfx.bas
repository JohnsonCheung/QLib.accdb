Attribute VB_Name = "MxAddIxPfx"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxAddIxPfx."


Function AddIxPfx(Ay, Optional b As EmIxCol = EiBeg0, Optional FmI&) As String()
If b = EiNoIx Then AddIxPfx = CvSy(Ay): Exit Function
Dim L, J&, N%
J = OffsetzEmBeg(b, FmI)
N = Len(CStr(UB(Ay) + J))
For Each L In Itr(Ay)
    PushI AddIxPfx, AlignR(J, N) & ": " & L
    J = J + 1
Next
End Function

Function AddIxPfxzLines(Lines, Optional b As EmIxCol = EiBeg0) As String()
AddIxPfxzLines = AddIxPfx(SplitCrLf(Lines), b)
End Function

