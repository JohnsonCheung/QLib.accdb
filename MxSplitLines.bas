Attribute VB_Name = "MxSplitLines"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxSplitLines."

Function SplitLineszDr(Dr, LinesColIx%) As Variant()
'Ret : a @Dy from @Dr & @LinesColIx.  The val of @Dr(@LinesColIx) is a Lines.  Split it as Ly.
'      for each Lin in Ly, build IDr with all val from Dr and Dr(LinesColIx) is Lin.
Dim IDr: IDr = Dr
Dim I:  For Each I In Itr(SplitCrLf(Dr(LinesColIx)))
    IDr(LinesColIx) = I
    Push SplitLineszDr, IDr
Next
End Function

Function SplitLineszDrs(A As Drs, LinesColNm$) As Drs
Dim Dy(): Dy = A.Dy
If Si(Dy) = 0 Then
    SplitLineszDrs = Drs(A.Fny, Dy)
    Exit Function
End If
Dim Ix%
    Ix = IxzAy(A.Fny, LinesColNm)
Dim O()
    Dim Dr
    For Each Dr In Dy
        PushAy Dy, SplitLineszDr(Dr, Ix)
    Next
SplitLineszDrs = Drs(A.Fny, O)
End Function
