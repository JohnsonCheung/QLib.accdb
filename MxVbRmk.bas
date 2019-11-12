Attribute VB_Name = "MxVbRmk"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbRmk."

Function IsLinVbRmk(L) As Boolean
IsLinVbRmk = FstChr(LTrim(L)) = "'"
End Function

Function VbRmk(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLinVbRmk(L) Then PushI VbRmk, L
Next
End Function

Function AeVbRmk(Ay) As String()
Dim L: For Each L In Itr(Ay)
    If Not IsLinVbRmk(L) Then PushI AeVbRmk, L
Next
End Function

Function EndTrimVbRmk(Src$()) As String()
Dim O$(): O = Src
Dim N%, J&: For J = UB(Src) To 0 Step -1
    Dim L$: L = LTrim(Src(J))
    Select Case True
    Case L = "", FstChr(LTrim(Src(J))) = "'": N = N + 1
    End Select
Next
If N > 0 Then
    ReDim Preserve O(UB(O) - N)
End If
EndTrimVbRmk = O
End Function

Function AwVbRmk(Ay) As String()
Dim L: For Each L In Itr(Ay)
    If IsLinVbRmk(L) Then PushI AwVbRmk, L
Next
End Function

