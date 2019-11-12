Attribute VB_Name = "MxFmtS12"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxFmtS12."

Function FmtS12s(A As S12s, Optional FF$ = "S1 S2", Optional IxCol As EmIxCol) As String()
If False Then
    FmtS12s = FmtDrszNoRdu(DrszS12s(A, FF), IxCol:=IxCol)
Else
    If A.N = 0 Then
        PushI FmtS12s, "(NoRec-S12s) (FF=" & FF & ")"
        Exit Function
    End If
    Dim D As Drs
    If Not HasLines(A) Then
               D = DrszFF(FF, DyzS12s(A))
         FmtS12s = FmtDrszNoRdu(D, Fmt:=EiSSFmt, IxCol:=IxCol)
:                  Exit Function
    End If
    Dim N1$, N2$:       AsgAp SyzSS(FF), N1, N2
    Dim S1$():     S1 = S1Ay(A)
    Dim S2$():     S2 = S2Ay(A)
    Dim W1%:       W1 = WdtzLinesAy(AddEleS(S1, N1))
    Dim W2%:       W2 = WdtzLinesAy(AddEleS(S2, N2))
    Dim W2Ay%(): W2Ay = IntAy(W1, W2)
    Dim SepL$:   SepL = LinzSep(W2Ay)
    Dim Tit$:     Tit = AlignDrWAsLin(Array(N1, N2), W2Ay)
    Dim M$():       M = FmtS12szW(A, W2Ay, SepL)              '  #Middle ! Middle part
    Dim W%:         W = Len(CStr(A.N))                        ' IxCol-Wdt
    Dim O$():       O = Sy(SepL, Tit, SepL, M)
                    O = AddIxCol(O, W, IxCol)               '          ! Add Ix col in front
    
    FmtS12s = O
End If
End Function

Function IxColLinPfx$(Fst2Chr$, IsIxAdd As Boolean, Sep$, Ix&, W%)
Dim O$
Select Case True
Case Fst2Chr = "|-":             O = Sep
Case Fst2Chr = "| " And IsIxAdd: O = "| " & Space(W + 1)
Case Fst2Chr = "| ":             O = "| " & Align(Ix, W) & " "
Case Else: Thw CSub, "Fst2Chr should [| ] or [|-]", "Fst2Chr", Fst2Chr
End Select
IxColLinPfx = O
End Function

Function AddIxCol(Fmt$(), W%, IxCol As EmIxCol, Optional IxBegI%) As String()
'@W   :Wdt #Ix-Col-Wdt#
'@Fmt :Ly                ! a formatted-Ly in format of
'                        ! Lin1: |----
'                        ! Lin2: | TTT
'                        ! Lin3: |----
'                        ! Lin4: | XXX
'Ret  : ! Add Ix column in front of @Fmt @@
Dim Ix&
    Select Case IxCol
    Case EmIxCol.EiBeg0
    Case EmIxCol.EiBeg1: Ix = 1
    Case EmIxCol.EiBegI: Ix = IxBegI
    Case Else:           AddIxCol = Fmt: Exit Function '==> Exit
    End Select


Dim S$: S = "|" & Dup("-", W + 2) ' Sep lin
Dim IsIxAdd As Boolean            ' Is-Ix-Added.
Dim F$                            ' Front str to be added in front of each line
Dim F2$ ' Fst 2 chr of each lin of @Fmt

PushI AddIxCol, S & Fmt(0)
PushI AddIxCol, "| " & AlignR("#", W) + " " & Fmt(1)

Dim J&: For J = 2 To UB(Fmt)
        F2 = Fst2Chr(Fmt(J))
        If F2 = "|-" Then IsIxAdd = False: Ix = Ix + 1
    F = IxColLinPfx(F2, IsIxAdd, S, Ix, W) 'What to add infront the a lin of @Fmt as an Ix col.
        If F2 = "| " And Not IsIxAdd Then IsIxAdd = True
        PushI AddIxCol, F & Fmt(J)
Next
End Function

Function FmtS12(A As S12, W2Ay%()) As String()
'@A    : the :S12.S1-S2 may both have lines.  Wrap them as @W2Ay.
'@W2Ay : S1-Wdt and S2-Wdt
'Ret   : Ly aft fmt @A @@
Dim Ly1$(), Ly2$()
    Ly1 = SplitCrLf(A.S1)
    Ly2 = SplitCrLf(A.S2)
          ResiMax Ly1, Ly2
    Ly1 = AlignAy(Ly1, W2Ay(0))
    Ly2 = AlignAy(Ly2, W2Ay(1))
Dim O$()
    Dim J%, Dr(): For J = 0 To UB(Ly1)
        Dr = Array(Ly1(J), Ly2(J))
:            PushI O, JnSpc(AlignDrzW(Dr, W2Ay))
    Next
FmtS12 = O
End Function

Function FmtS12szW(A As S12s, W2Ay%(), SepL$) As String()
'Ret :  #Middle ! Middle part @@
Dim J&: For J = 0 To A.N - 1
    PushIAy FmtS12szW, FmtS12(A.Ay(J), W2Ay)
    PushI FmtS12szW, SepL
Next
'Insp "QVb_S1S2_Fmt.FmtS12szW", "Inspect", "Oup(FmtS12szW) A W2Ay SepL", FmtS12szW, FmtS12s(A), W2Ay, SepL: Stop
End Function

Function HasLines(A As S12s) As Boolean
Dim J&
HasLines = True
For J = 0 To A.N - 1
    With A.Ay(J)
        If IsLines(.S1) Then Exit Function
        If IsLines(.S2) Then Exit Function
    End With
Next
HasLines = False
End Function

Sub Z_FmtS12s()
Dim A As S12s, FF$, Pseg$
'GoSub T0
'GoSub T1
GoSub T2
'GoSub T3
Exit Sub
T3:
    FF = "AA BB"
    Pseg = "Z_FmtS12s\Cas3"
    A = S12szRes("S12s.Txt", Pseg & "\Inp")
    Ept = Resl("Ept", Pseg)
    GoTo Tst
T0:
    FF = "AA BB"
    A = AddS12(S12("A", "B"), S12("AA", "B"))
    GoTo Tst
T1:
    FF = "AA BB"
    A = SampS12s
    GoTo Tst
T2:
    FF = "AA BB"
    A = SampS12s
    Brw FmtS12s(A, FF)
    Stop
    GoTo Tst
Tst:
    Act = FmtS12s(A, FF)
    C
    Return
End Sub
