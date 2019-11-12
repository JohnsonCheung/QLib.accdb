Attribute VB_Name = "MxULin"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxULin."

Function ULinDbl$(Lin, Optional IncLin As Boolean)
ULinDbl = IIf(IncLin, Lin & vbCrLf, "") & String(Len(Lin), "=")
End Function

Function PushMsgULinDbl(O$(), M$)
Push O, M
Push O, ULinDbl(M)
End Function

Function PushULin(O$())
Push O, ULin(LasEle(O))
End Function

Function PushULinDbl(O$())
Push O, ULinDbl(LasEle(O))
End Function

Function ULinzLines$(Lines$, Optional ULinChr$ = "-")
ULinzLines = Lines & vbCrLf & Dup("-", WdtzLines(Lines))
End Function

Function PushMsgULin(O$(), M$)
Push O, M
Push O, ULin(M)
End Function

Function ULin$(S$, Optional ULChr$ = "-")
ULin = Dup(ULChr, Len(S))
End Function
