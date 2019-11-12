Attribute VB_Name = "MxUtfSig"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxUtfSig."
Public Const Utf8Sig$ = "﻿"

Function RmvUtf8Sig$(S$)
RmvUtf8Sig = RmvPfx(S, Utf8Sig)
End Function

Function Z_HasUtfSig8()
Dim F$: F = LineszFt(ResFcsv("DoMthP"))
Debug.Assert HasUtf8Sig(F)
End Function

Function HasUtf8Sig(S$) As Boolean
HasUtf8Sig = HasPfx(S, Utf8Sig, vbBinaryCompare)
End Function

