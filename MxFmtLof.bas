Attribute VB_Name = "MxFmtLof"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxFmtLof."
Public Const VdtLofAlivv$ = "Left Right Centre"
Public Const VdtLofBdrvv$ = "Left Right Both"
':Lof: :Ly #ListObject-Formatter# ! Each line is Ly with T1 LofLofT1nn"
':FldLikss: :Likss #Fld-Lik-SS# ! A :SS to expand a given Fny
':Ali:   :
Sub FmtLo(Lo As ListObject, Lof$())
Dim F$(): F = FnyzLo(Lo)
ThwIf_Er EoLof(Lof, F), CSub
Dim D As Dictionary: Set D = DiT1qLyItr(Lof, VdtLofT1ss)
Dim L
For Each L In D("Ali"): SetLoAli Lo, L, F: Next
For Each L In D("Bdr"): SetLoBdr Lo, L, F: Next
For Each L In D("Bet"): SetLoBet Lo, L:    Next
For Each L In D("Cor"): SetLoCor Lo, L, F: Next
For Each L In D("Fml"): SetLoFml Lo, L:    Next
For Each L In D("Fmt"): SetLoFmt Lo, L, F: Next
For Each L In D("Lvl"): SetLoLvl Lo, L, F: Next
For Each L In D("Tot"): SetLoTot Lo, L, F: Next
For Each L In D("Wdt"): SetLoWdt Lo, L, F: Next
SetLoTit Lo, LyzLyItr(D("Tit"))
SetLon Lo, T2(FstElezT1(Lof, "Nm"))
For Each L In D("Lbl"): SetLoLbl Lo, L: Next ' Must run Last
End Sub

'Ali -----------------------------------------------------------
Sub SetLoAli(L As ListObject, LinOf_Ali_FldLikss, Fny$())
'Fm LinOf_Ali_FldLikss : T1 is :Ali: Rst is FldLikss.  :Ali: is 'Left | Right | Center'
'Ret                   : align those col as stated in @FldLikss and in @LoFny as @Ali
Dim Ali$:     Ali = T1(LinOf_Ali_FldLikss)
Dim Fny1$(): Fny1 = AwLikss(Fny, RmvT1(LinOf_Ali_FldLikss))
Dim H As XlHAlign: H = HAlign(Ali)
Dim F: For Each F In Itr(Fny1)
    EntColRgzLc(L, F).HorizontalAlignment = H
Next
End Sub

Function HAlign(Ali$) As XlHAlign
Select Case Ali
Case "Left": HAlign = xlHAlignLeft
Case "Right": HAlign = xlHAlignRight
Case "Center": HAlign = xlHAlignCenter
Case Else: Inf CSub, "Invalid Ali", "Valid Ali", VdtLofAlivv: Exit Function
End Select
End Function
Function FnyzT1FldLikss(Fny$(), LinOf_T1_FldLikss) As String()
FnyzT1FldLikss = AwLikss(Fny, RmvT1(LinOf_T1_FldLikss))
End Function
Sub SetLoBdr(L As ListObject, LinOf_LRBoth_FldLikss, Fny$())
'#  A.Fny                 :
'Fm LinOf_LRBoth_FldLikss : T1 is :Ali: Rst is FldLikss.  :LRBoth: is 'Left | Right | Both'
'Ret                   : align those col as stated in @FldLikss and in @LoFny as @Ali

Dim LRBoth$: LRBoth = T1(LinOf_LRBoth_FldLikss)
Dim Fny1$():   Fny1 = FnyzT1FldLikss(Fny, LinOf_LRBoth_FldLikss)
Dim IsLeft As Boolean: IsLeft = HasEle(SyzSS("Left Both"), LRBoth)
Dim IsRight As Boolean: IsRight = HasEle(SyzSS("Right Both"), LRBoth)
Dim F: For Each F In Itr(Fny)
    If IsLeft Then BdrRgAy ColRgAy(L, Fny), xlEdgeLeft
    If IsRight Then BdrRgAy ColRgAy(L, Fny), xlEdgeRight
Next
End Sub

Sub SetLoBet(L As ListObject, LinOf_Sum_Fm_To)
Dim FSum$, FFm$, FTo$: AsgTTRst LinOf_Sum_Fm_To, FSum, FFm, FTo
EntColRgzLc(L, FSum).Formula = FmtQQ("=Sum([?]:[?])", FFm, FTo)
End Sub

Sub SetLoCor(L As ListObject, LinOf_Cor_FldLikss, Fny$())
Dim Cor$: Cor = T1(LinOf_Cor_FldLikss)
Dim Fny1$(): Fny1 = FnyzT1FldLikss(Fny, LinOf_Cor_FldLikss)
Dim C&: C = Colr(Cor)
Dim F: For Each F In Itr(Fny1)
    EntColRgzLc(L, F).Color = C
Next
End Sub

Sub SetLoFml(L As ListObject, LinOf_Fld_Fml)
Dim F$, Fml$: AsgTRst LinOf_Fld_Fml, F, Fml
EntColRgzLc(L, F).Formula = Fml
End Sub

Sub SetLoFmt(L As ListObject, LinOf_Fmt_FldLikss, Fny$())
Dim Fmt$: Fmt = T1(LinOf_Fmt_FldLikss)
Dim Fny1$(): Fny1 = FnyzT1FldLikss(Fny, LinOf_Fmt_FldLikss)
Dim F: For Each F In Itr(Fny1)
    EntColRgzLc(L, F).NumberFormat = Fmt
Next
End Sub

Sub SetLoLbl(L As ListObject, LinOf_Fld_Lbl)
'Ret: Must run after forumla & Between
Dim Fld$, Lbl$:    AsgTRst LinOf_Fld_Lbl, Fld, Lbl
Dim R1 As Range
Dim R2 As Range
Set R1 = LoHdrCell(L, Fld)
Set R2 = CellAbove(R1)
SwapCellVal R1, R2
End Sub

Sub AsgT1Fny(LinOf_T1_FldLikss, Fny$(), OT1, OFny$())
OFny = FnyzT1FldLikss(Fny, LinOf_T1_FldLikss)
OT1 = T1(LinOf_T1_FldLikss)
End Sub

Sub SetLoLvl(L As ListObject, LinOf_Lvl_FldLikss, Fny$())
Dim XFny$(), XLvl As Byte: AsgT1Fny LinOf_Lvl_FldLikss, Fny, XLvl, XFny
Dim F: For Each F In Itr(XFny)
    EntColRgzLc(L, F).OutlineLevel = XLvl
Next
End Sub

Sub SetLoTot(L As ListObject, LinOf_SACnt_FldLikss, Fny$())
Dim XSumAvgCnt$, XFny$():              AsgT1Fny LinOf_SACnt_FldLikss, Fny, XSumAvgCnt, XFny
Dim T As XlTotalsCalculation: T = XTotCalc(XSumAvgCnt)
Dim F: For Each F In Itr(XFny)
    L.ListColumns(F).Total = T
Next
End Sub

Function XTotCalc(SumAvgCnt$) As XlTotalsCalculation
'Fm SACnt : "Sum | Avg | Cnt" @@
Dim O As XlTotalsCalculation
Select Case SumAvgCnt
Case "Sum": O = xlTotalsCalculationSum
Case "Avg": O = xlTotalsCalculationAverage
Case "Cnt": O = xlTotalsCalculationCount
Case Else: Inf CSub, "Invalid TotCalcStr", "TotCalcStr Valid-TotCalcStr", SumAvgCnt, "Sum Avg Cnt": Exit Function
End Select
XTotCalc = O
End Function

'Wdt----------------------------------------------------------
Sub SetLoWdt(L As ListObject, LinOf_Wdt_FldLikss, Fny$())
Dim Wdt$, Fny1$(): AsgT1Fny LinOf_Wdt_FldLikss, Fny, Wdt, Fny1
Dim W%: W = Wdt
If Not IsBet(W, 5, 200) Then Inf CSub, "Invalid Wdt (should between 5 200)", "Wdt Lin", W, LinOf_Wdt_FldLikss: Exit Sub
Dim F: For Each F In Itr(Fny)
    RgzLc(L, F).EntireColumn.Width = W
Next
End Sub

'Tst-------------------------------------------------------------
Sub Z_FmtLo()
Dim Lo As ListObject, Fmtr() As String 'Lofr
'------------
Set Lo = SampLo
Fmtr = SampLof
GoSub Tst
Exit Sub
Tst:
    FmtLo Lo, Fmtr
    Return
End Sub

Sub Z_SetBdr()
Dim Lin$, Lo As ListObject, Fny$()
'--
Set Lo = SampLo
Fny = FnyzLo(Lo)
'--
GoSub T1
GoSub T2
Exit Sub
T1: Lin = "Left A B C": GoTo Tst
T2: Lin = "Left D E F": GoTo Tst
T3: Lin = "Right A B C": GoTo Tst
T4: Lin = "Center A B C": GoTo Tst
Tst:
    SetLoBdr Lo, L, Fny     '<=='
    Stop
    Return
End Sub


'Fun===========================================================================
Function LoHdrCell(L As ListObject, C) As Range
Set LoHdrCell = A1zRg(CellAbove(L.ListColumns(C).Range))
End Function

Sub FmtLoBStd(b As Workbook)
Dim S As Worksheet
For Each S In b.Sheets
    FmtLoSStd S
Next
End Sub
Sub FmtLoStd(L As ListObject)
FmtLo L, StdLof
End Sub

Sub FmtLoSStd(S As Worksheet)
Dim L As ListObject: For Each L In S.ListObjects
    FmtLo L, StdLof
Next
End Sub

Property Get StdLof() As String()

End Property

Sub AddLoFml(L As ListObject, ColNm$, Fml$)
Dim O As ListColumn
Set O = L.ListColumns.Add
O.Name = ColNm
O.DataBodyRange.Formula = Fml
End Sub

Sub AutoFit(L As ListObject, Optional MaxW = 100)
Dim C As Range: Set C = LoAllEntCol(L)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = EntRgC(C, J)
   If EntC.ColumnWidth > MaxW Then EntC.ColumnWidth = MaxW
Next
End Sub

Sub SetLonUgTbl(L As ListObject)
SetLon L, L.QueryTable.CommandText
End Sub

Sub SetLon(L As ListObject, Lon$)
If Lon <> "" Then
    If Not HasLo(WszLo(L), Lon) Then
        L.Name = Lon
    Else
        Inf CSub, "Lo"
    End If
End If
End Sub

Sub SetLcWrp(L As ListObject, ColTermLin$, Optional Wrp As Boolean)
Dim C: For Each C In TermAy(ColTermLin)
    SetLcWrp_ L, C, Wrp
Next
End Sub

Sub SetLcWrp_(L As ListObject, C, Optional Wrp As Boolean)
L.ListColumns(C).DataBodyRange.WrapText = Wrp
End Sub

Sub SetLcWdt(L As ListObject, ColTermLin$, W)
Dim C: For Each C In TermAy(ColTermLin)
    SetLcWdt_ L, C, W
Next
End Sub

Sub SetLcWdt_(L As ListObject, C, W)
EntColzLc(L, C).ColumnWidth = W
End Sub

Sub SetLoTotLnk(L As ListObject, C)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = L.ListColumns(C).DataBodyRange
Set Ws = WszRg(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, R.Rows.Count + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

Sub SetLoTit(L As ListObject, TitLy$())
Dim Sq(), R As Range
    Sq = XTitSq(TitLy, FnyzLo(L)): If Si(Sq) = 0 Then Exit Sub
    Set R = XTitAt(L, UBound(Sq(), 1))
Set R = RgzSq(Sq(), R)
XMgeTit R
BdrInside R
BdrAround R
End Sub

Sub XMgeTit(TitRg As Range)
Dim J%
For J = 1 To TitRg.Rows.Count
    XMgeTitH RgR(TitRg, J)
Next
For J = 1 To TitRg.Columns.Count
    XMgeTitV RgC(TitRg, J)
Next
End Sub

Sub XMgeTitH(TitRg As Range)
TitRg.Application.DisplayAlerts = False
Dim J%, C1%, C2%, V, LasV
LasV = RgRC(TitRg, 1, 1).Value
C1 = 1
For J = 2 To TitRg.Columns.Count
    V = RgRC(TitRg, 1, J).Value
    If V <> LasV Then
        C2 = J - 1
        If Not IsEmpty(LasV) Then
            RgRCC(TitRg, 1, C1, C2).MergeCells = True
        End If
        C1 = J
        LasV = V
    End If
Next
TitRg.Application.DisplayAlerts = True
End Sub

Sub XMgeTitV(A As Range)
Dim J%
For J = A.Rows.Count To 2 Step -1
    MgeCellAbove RgRC(A, J, 1)
Next
End Sub

Function XTitAt(Lo As ListObject, NTitRow%) As Range
Set XTitAt = RgRC(Lo.DataBodyRange, 0 - NTitRow, 1)
End Function

Function XTitSq(TitLy$(), LoFny$()) As Variant()
Dim Fny$()
Dim Col()
    Dim F$, I, Tit$
    For Each I In Fny
        F = I
        Tit = FstElezRmvT1(TitLy, F)
        If Tit = "" Then
            PushI Col, Sy(F)
        Else
            PushI Col, AmTrim(SplitVBar(Tit))
        End If
    Next
XTitSq = Transpose(SqzDy(Col))
End Function

Sub Z_XTitSq()
Dim TitLy$(), Fny$()
'----
Dim A$(), Act(), Ept()
'TitLy
    Erase A
    Push A, "A A1 | A2 11 "
    Push A, "B B1 | B2 | B3"
    Push A, "C C1"
    Push A, "E E1"
    TitLy = A
Fny = SyzSS("A B C D E")
Ept = XTitSq(TitLy, Fny)
    SetSqr Ept, 1, SyzSS("A1 B1 C1 D E1")
    SetSqr Ept, 2, Array("A2 11", "B2")
    SetSqr Ept, 3, Array(Empty, "B3")
GoSub Tst
Exit Sub
'---
'TitLy
    Erase A
    Push A, "ksdf | skdfj  |skldf jf"
    Push A, "skldf|sdkfl|lskdf|slkdfj"
    Push A, "askdfj|sldkf"
    Push A, "fskldf"
    TitLy = A
BrwSq XTitSq(TitLy, Fny)

Exit Sub
Tst:
    Act = XTitSq(TitLy, Fny)
    Ass IsEqSq(Act, Ept)
    Return
End Sub

Sub BrwSampLof()
Brw FmtLof(SampLof)
End Sub

Function FmtLof(Lof$()) As String()
Dim mAmT1$():    mAmT1 = TermAy(VdtLofT1ss)
Dim O$()
    Dim T$, I
    For Each I In mAmT1
        T = I
        PushIAy O, AwT1(Lof, T)
    Next
    Dim M$(): M = SyeT1Sy(Lof, mAmT1)
    If Si(M) > 0 Then
        PushI O, FmtQQ("# Error: in not AmT1(?)", TLin(mAmT1))
        PushIAy O, M
    End If
FmtLof = AlignLyz2T(O)
End Function
