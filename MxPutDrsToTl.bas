Attribute VB_Name = "MxPutDrsToTl"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPutDrsToTl."
Sub Z_PutDrsToLo()
Dim Lo As ListObject, D As Drs
GoSub Z
Exit Sub

Z:
Set Lo = CrtLozDrs(SampDrs1, NewA1)
PutDrsToLo SampDrs2, Lo
Stop
'PutDrsToLo SampDrs3, Lo
Stop
ClsWbNoSav WbzLo(Lo)
Return
End Sub
Sub PutDrsToLo(D As Drs, Lo As ListObject)
ClrLo Lo
PutDyToLo SelDrsAlwEzFny(D, FnyzLo(Lo)).Dy, Lo
End Sub
Sub PutDyToLo(Dy(), Lo As ListObject)
RgzDy Dy, RgRC(Lo.Range, 2, 1)
End Sub

