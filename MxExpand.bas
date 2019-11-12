Attribute VB_Name = "MxExpand"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxExpand."

Function Expand(QVblTp$, Seed$()) As String()
Dim Tp$, ISeed
Tp = RplVBar(QVblTp)
For Each ISeed In Itr(Seed)
    PushI Expand, RplQ(Tp, ISeed)
Next
End Function

Function Expandss(QVblTp$, Seedss$) As String()
Expandss = Expand(QVblTp, SyzSS(Seedss))
End Function
Sub Z_Expandss()
Dim QVblTp$, Seed$()
Z:
    Erase XX
    X "Sub Push?(O() As ?, M As ?)"
    X "Dim N&"
    X "N = ?Si(O)"
    X "ReDim Preserve O(N)"
    X "O(N) = M"
    X "End Sub"
    X ""
    X "Function ?Si&(A() As ?)"
    X "On Error Resume Next"
    X "?Si = Ubound(A) + 1"
    X "End Function"
    X ""
    X ""
    QVblTp = JnVBar(XX)
    Erase XX
    Brw Expandss(QVblTp, "S12 XX")
T0:
    QVblTp = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
    Seed = SyzSS("Xws Xwb Xfx Xrg")
    Erase XX
    X ""
    X ""
    Ept = JnCrLf(XX)
    Erase XX
    GoTo Tst
Tst:
    Act = Expand(QVblTp, Seed)
    C
    Return
End Sub

