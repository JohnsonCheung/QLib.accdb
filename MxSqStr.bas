Attribute VB_Name = "MxSqStr"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxSqStr."
Function SqStr$(Sq())
':SqStr: :S  ! it is Lines of CellStr lin.
'            ! #CellStr-Lin is :CellStr separaterd by vbTab with ending vbTab.
'            ! the fst CellStr-Lin will not trim the ending vbTab, because it is used to determine how many col.
'            ! if any non-Lin1-CellStr-Lin has more fld than lin1-CellStr-lin-fld, the extra fld are ignored and inf (this is done in %SqS%
'            ! the reverse fun is %SqzS @@
Dim L$(), O$()
Dim UC%: UC = UBound(Sq, 1)
Dim R&: For R = 1 To UBound(Sq, 1)
    ReDim L(UC)
    Dim C&: For C = 1 To UBound(Sq, 2)
        L(C - 1) = CellStr(Sq(R, C))
    Next
    PushI O, JnTab(L)
Next
SqStr = JnCrLf(O)
End Function

Function SqzS(SqStr$) As Variant()
'Ret : :Sq from :SqStr
Dim Ry$(): Ry = SplitCrLf(SqStr): If Si(Ry) = 0 Then Exit Function
Dim NR&: NR = Si(Ry)
Dim R1$: R1 = Ry(0)
Dim NC%: NC = Si(SplitTab(R1))
Dim O(): ReDim O(1 To NR, 1 To NC)
Dim IR&, IC%
Dim R: For Each R In Ry
    IR = IR + 1
    Dim C: For Each C In SplitTab(R)
        IC = IC + 1
        If IC > NC Then Exit For ' ign the extra fld, if it has more fld then lin1-fld-cnt
        O(IR, IC) = VzLStr(C)
    Next
Next
End Function

Function VzLStr(LStr)
'Ret : ! a val (Str|Dbl|Bool|Dte|Empty) fm @LStr.
'      ! If fst letter is
'      !   ['] is a str wi \r\n\t
'      !   [D] is a str of date, if cannot convert to date, ret empty and debug.print msg.
'      !   [T] is true
'      !   [F] is false
'      !   rest is dbl, if cannot convert to dbl, ret empty and debug.print msg @@
Dim F$: F = FstChr(LStr)
Dim O$
Select Case F
Case "'": O = UnSlashCrLfTab(RmvFstChr(LStr))
Case "T": O = True
Case "F": O = False
Case "D": O = CvDte(RmvFstChr(LStr))
Case ""
Case Else: O = CvDbl(RmvFstChr(LStr))
End Select
VzLStr = O
End Function

Function LStr$(V, Optional Fun$)
':LStr: :S #Letter-Str# ! A str wi fst letter can-determine the str can converted to what value.
Dim T$: T = TypeName(V)
Dim O$
Select Case T
Case "String": O = "'" & SlashCrLfTab(V)
Case "Boolean": O = IIf(V, "T", "F")
Case "Integer", "Single", "Double", "Currency", "Long": O = V
Case "Date": O = "D" & V
Case Else: If Fun <> "" Then Inf CSub, "Val-of-TypeName[" & T & "] cannot cv to :LStr"
End Select
LStr = O
End Function



