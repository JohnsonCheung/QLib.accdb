Attribute VB_Name = "MxSrcc"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrcc."
':Srcc: :Src #Src-Cleaned# ! Src without blank/rmk-Line
Function SrccP() As String()
SrccP = SrcczP(CPj)
End Function

Function SrcHasSngDblQP() As String()
SrcHasSngDblQP = SrcHasSngDblQ(SrcczP(CPj))
End Function

Sub Z_SrcHasSngDblQP()
Brw SrcHasSngDblQP
End Sub

Function SrcHasSngDblQ(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If HasSngDblQ(L) Then
        PushI SrcHasSngDblQ, L
    End If
Next
End Function

Function HasSngDblQ(S) As Boolean
If HasSngQ(S) Then
    If HasDblQ(S) Then
        HasSngDblQ = True
    End If
End If
End Function

Function SrcczM(M As CodeModule) As String()
SrcczM = Srcc(Src(M))
End Function

Function SrcczP(P As VBProject) As String()
SrcczP = Srcc(SrczP(P))
End Function

Function Srccs(Srcc$()) As String()
':Srccs: :Src #Cleared-Empty-Lin/Rmk-And-Clear-String# ! All string-cxt quoted by DblQ are removed
Dim L: For Each L In Itr(Srcc)
    PushI Srccs, SrccsLin(L)
Next
End Function

Function SrccsLin$(L)
Dim O$: O = RplDblSpc(L)
Dim J%
X:
J = J + 1: If J > 10000 Then ThwLoopingTooMuch CSub
Dim P1&: P1 = InStr(O, vbDblQ): If P1 = 0 Then SrccsLin = L: Exit Function
Dim P2&: P2 = InStr(P1 + 1, O, vbDblQ): If P2 = 0 Then Stop
O = Left(O, P1 - 1) & Mid(O, P2 + 1)
GoTo X
End Function

Function Srcc(Src$()) As String()
':Srcc: :Src #Cleared-Empty-Lin/Rmk-Src# ! all empty-lin and rmk-lin are removed.
Dim L: For Each L In Itr(Src)
    If IsLinCd(L) Then PushI Srcc, L
Next
End Function

Function IsLinNCd(L) As Boolean
IsLinNCd = Not IsLinCd(L)
End Function
Function IsLinCd(L) As Boolean
Dim A$: A = Trim(L)
If A = "" Then Exit Function
If FstChr(A) = "'" Then Exit Function
IsLinCd = True
End Function

Function IsLinNonOpt(Lin) As Boolean
If Not IsLinCd(Lin) Then Exit Function
If HasPfx(Lin, "Option") Then Exit Function
IsLinNonOpt = True
End Function
