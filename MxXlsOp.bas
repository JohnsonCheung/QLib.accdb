Attribute VB_Name = "MxXlsOp"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsOp."
Enum EmWsPos
    EiEnd
    EiBeg
    EiRfWs
End Enum

Sub Brw_Fx(Fx)
If HasFfn(Fx) Then Debug.Print "No Fx:" & Fx
ShwWb WbzFx(Fx)
End Sub

Sub CrtFx(Fx)
WbSavAs(NewWb, Fx).Close
End Sub

Function EnsFx$(Fx)
'Ret : crt an emp @Fx if not exist @@
If NoFfn(Fx) Then CrtFx Fx
EnsFx = Fx
End Function

Sub Ins_Row_At(OCell As Range, Optional N = 1)
'Fm OCell : due to xls ins row will the rg used to ins row, so @OCell will be rest to A1-of-original-@OCell.
'Ret      : ins @N emp rows below the fst row of @OCell
Dim Ws As Worksheet: Set Ws = WszRg(OCell)
Dim R&:                   R = OCell.Row
Dim C%:                   C = OCell.Column
Dim OIns:                     RgCRR(OCell, 1, 1, N).EntireRow.Insert ' @OCell will be changed after insert  ! <==
                  Set OCell = WsRC(Ws, R, C)                         ' reset @OCell to A1-of-orignal-@OCell ! <==
End Sub

Sub RmvEmpRow(OInRg As Range)
'Fm OInRg : ! the rect may have emp row.  Aft rmv, @OInRg will chg to pointing to the remaining rect.
'Ret    : ! ##Rmv_NRow_At-Wht:Emp-Fm:OInRg Rmv all emp row ept keeping one emp row.  Rows rmv is ent row.
'         ! All non emp row will keep no chg and may shf up.
'         ! set the @OInRg pointing to
Dim Ws As Worksheet: Set Ws = WszRg(OInRg)     '                        ! <--
Dim R&:                   R = OInRg.Row        ' A1-Rno-of-@OInRg         ! <--
Dim C%:                   C = OInRg.Column     ' A1-Cno-of-@OInRg         ! <--
Dim NR%:                 NR = OInRg.Rows.Count
Dim C2%:                 C2 = NColzRg(OInRg)     ' N-Columns-of-@OInRg      ! <--
Dim Sq():                Sq = SqzRgNo(OInRg)
Dim RRmv&():           RRmv = RxyzEmpRow(Sq) ' Rny to be rmv '        ! <--
Dim NRLeft&:         NRLeft = NR - Si(R)     ' R2-of-the-remaining-rg
Dim R2&:                 R2 = R + NRLeft - 1 '                        ! <--
'=======================================================================================================================
EntRows(Ws, RRmv).Remove
Set OInRg = WsRCRC(Ws, R, C, R2, C2)
End Sub

Function EntRowsStr$(Rny&())
Dim O$()
Dim Rno: For Each Rno In Rny
    PushI O, Rno & ":" & Rno
Next
EntRowsStr = JnComma(O)
End Function

Function EntRows(Ws As Worksheet, Rny&()) As Range
Set EntRows = Ws.Range(EntRowsStr(Rny))
End Function

Sub Rmv_NRow_At(OAt As Range, Optional N = 1)
'Fm OAt : due the xls rmv row, the range used to rmv row will be changed, so @OAt will be rest to A1-of-original-@OAt.
'Ret      : rmv the @N-rows fm the fst of @OAt.
Dim Bef$: Bef = A1zRg(OAt).Address    ' Original A1-of-@OAt
RgCRR(OAt, 1, 1, N).EntireRow.Insert ' @OAt will be changed after insert
Set OAt = RgRC(OAt, -N + 1, 1)     ' reset @OAt to A1-of-orignal-@OAt
Dim Aft$: Aft = OAt.Address          ' Aft rmv row, the adr-of-@OAt
If Bef <> Aft Then Stop
End Sub

Function OpnFx(Fx) As Workbook
ThwIf_NoFfn Fx, CSub
Set OpnFx = Xls.Workbooks.Open(Fx)
End Function

Sub OpnFcsv(Fcsv)
Xls.Workbooks.OpenText Fcsv
End Sub

Sub ClrWsNm(Ws As Worksheet)
Dim N
For Each N In Itn(Ws.Names)
    Ws.Names(N).Delete
Next
End Sub

Sub XEnsHyp(Cell As Range, W$())
'Fm W : Ws ny
'Ret : Ens the @Cell is pointing of one of @W if the @cell val is in @W else rmv the hyp lnk @@
Dim V:  V = Cell.Value
Dim Wsn$: If IsStr(V) Then If HasEle(W, V) Then Wsn = V  ' The wsn pointed by the @Cell.Value, if the @Cell.Value is str and in @W
Dim A$                                                   ' The adr of A1 of %Wsn
    If Wsn <> "" Then A = WszWb(WbzRg(Cell), Wsn).Range("A1").Address(External:=True)
Dim L As Hyperlink: For Each L In Cell.Hyperlinks
    If L.SubAddress <> A Then Cell.Hyperlinks.Delete
Next
'Need to add if #A<>"" and @Cell hyp lnk cnt = 0
With Cell.Hyperlinks
    If A <> "" And .Count = 0 Then
        .Add Anchor:=Cell, Address:="", SubAddress:=A
    End If
End With
End Sub
Sub Z_EnsHyp()
Dim Rg As Range
GoSub Crt
GoSub T0
GoSub Rmv
Exit Sub
T0:
    Set Rg = CWs.Range("A1:A2")
    EnsHyp Rg
    'Stop
    Return
Crt:
    Dim Wb As Workbook: Set Wb = NewWb("A")
    AddWs Wb, "A 1"
    AddWs Wb, "A 2"
    WszWb(Wb, "A").Activate
    CWs.Range("A1").Value = "A 1"
    CWs.Range("A2").Value = "A 2"
    ShwWb Wb
    Return
Rmv:
    Wb.Close False
    Return
End Sub
Sub EnsHyp(Rg As Range)
'Ret: Ens the hyp lnk in @Rg is pointing ws A1 if the @Rg val is a ws name. @@
Dim W$(): W = Wny(WbzRg(Rg))   ' Ws ny
Dim Cell As Range: For Each Cell In Rg
    XEnsHyp Cell, W 'Ens the @Cell is pointing of one of @W if the @cell val is in @W else rmv the hyp lnk
Next
End Sub

Sub RmvWsIf(Fx, Wsn$)
'Ret : Ret @Wsn in @Fx if exists @@
If HasFxw(Fx, Wsn) Then
   Dim b As Workbook: Set b = WbzFx(Fx)
   WszWb(b, Wsn).Delete
   SavWb b
   ClsWbNoSav b
End If
End Sub

Function LozAyH(Ay, Wb As Workbook, Optional Wsn$, Optional Lon$) As ListObject
Set LozAyH = CrtLo(RgzSq(SqH(Ay), A1zWb(Wb, Wsn)), Lon)
End Function

Sub Z_SetWsCdNm()
Dim A As Worksheet: Set A = NewWs
SetWsCdNm A, "XX"
ShwWs A
Stop
End Sub
Sub MgeBottomCell(VBar As Range)
Ass IsRgVBar(VBar)
Dim R2: R2 = VBar.Rows.Count
Dim R1
    Dim Fnd As Boolean
    For R1 = R2 To 1 Step -1
        If Not IsEmpty(RgRC(VBar, R1, 1)) Then Fnd = True: GoTo Nxt
    Next
Nxt:
    If Not Fnd Then Stop
If R2 = R1 Then Exit Sub
Dim R As Range: Set R = RgCRR(VBar, 1, R1, R2)
R.Merge
R.VerticalAlignment = XlVAlign.xlVAlignTop
End Sub

Sub SetWsCdNm(A As Worksheet, CdNm$)
CmpzWs(A).Name = CdNm
End Sub

Sub SetWsCdNmAndLon(A As Worksheet, Nm$)
CmpzWs(A).Name = Nm
SetLon FstLo(A), Nm
End Sub
Function PutDbtAtzByWc(Db As Database, T, At As Range) As Range

End Function


Function CvCmt(A) As Comment
Set CvCmt = A
End Function

Function IsCell(A As Range) As Boolean
If A.Rows.Count > 1 Then Exit Function
If A.Columns.Count > 1 Then Exit Function
IsCell = True
End Function

Function HasCell(A As Range, Cell As Range) As Boolean
If Not IsCell(Cell) Then Exit Function
If Not IsBet(Cell.Row, A.Row, NRowzRg(A)) Then Exit Function
If Not IsBet(Cell.Column, A.Column, NColzRg(A)) Then Exit Function
HasCell = True
End Function
Function CmtAyzRg(A As Range) As Comment()
Dim C As Comment: For Each C In WszRg(A).Comments
    Dim CmtRg As Range: Set CmtRg = C.Parent
    If HasCell(A, CmtRg) Then PushObj CmtAyzRg, C
Next
End Function
Sub DltCmtzRg(A As Range)
Dim Cmt() As Comment: Cmt = CmtAyzRg(A)
Dim I: For Each I In Itr(Cmt)
    CvCmt(I).Delete
Next
End Sub



Sub MgeRg(A As Range)
A.MergeCells = True
A.HorizontalAlignment = XlHAlign.xlHAlignCenter
A.VerticalAlignment = XlVAlign.xlVAlignCenter
End Sub



Sub ThwWbMisOupNy(A As Workbook, OupNy$())
Dim O$(), N$, b$(), Wny$()
Wny = WsCdNy(A)
O = AyMinus(AmAddPfx(OupNy, "WsO"), Wny)
If Si(O) > 0 Then
    N = "OupNy":  b = OupNy:  GoSub Dmp
    N = "WbCdNy": b = Wny: GoSub Dmp
    N = "Mssing": b = O:      GoSub Dmp
    Stop
    Exit Sub
End If
Exit Sub
Dmp:
Debug.Print ULin(N)
Debug.Print N
Debug.Print ULin(N)
DmpAy b
Return
End Sub

Sub ClsAllWbNoSav()
Dim X As Excel.Application: Set X = Xls
While X.Workbooks.Count > 0
    ClsWbNoSav X.Workbooks(1)
Wend
End Sub
Sub ClsLasWbNoSav()
ClsWbNoSav LasWb
End Sub
Sub ClsWbNoSav(A As Workbook)
A.Close False
End Sub


Function MaxiWb(A As Workbook) As Workbook
A.Application.WindowState = xlMaximized
Set MaxiWb = A
End Function

Sub QuitWb(A As Workbook)
QuitXls A.Application
End Sub

Function SavWbCsv(A As Workbook, Fcsv$) As Workbook
DltFfnIf Fcsv
A.Application.DisplayAlerts = False
A.SaveAs Fcsv, XlFileFormat.xlCSV
A.Application.DisplayAlerts = True
Set SavWbCsv = A
End Function

Function SavWb(A As Workbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.Save
A.Application.DisplayAlerts = Y
Set SavWb = A
End Function

Function WbSavAs(A As Workbook, Fx, Optional Fmt As XlFileFormat = xlOpenXMLWorkbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.SaveAs Fx, Fmt
A.Application.DisplayAlerts = Y
Set WbSavAs = A
End Function

Sub SetWcFcsv(A As Workbook, Fcsv$)
'Set first Wb TextConnection to Fcsv if any
Dim T As TextConnection: Set T = TxtWc(A)
Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub

Function HasWsCd(WsCdn$, Optional IsInf As Boolean) As Boolean
Dim Ws As Worksheet
For Each Ws In CWb.Sheets
    If Ws.CodeName = WsCdn Then HasWsCd = True: Exit Function
Next
If IsInf Then
    MsgBox RplVBar("No worksheet code||" & WsCdn), vbCritical
End If
End Function
Function HasWs(A As Workbook, WsIx) As Boolean
If IsNumeric(WsIx) Then
    HasWs = IsBet(WsIx, 1, A.Sheets.Count)
    Exit Function
End If
Dim Ws As Worksheet
For Each Ws In A.Worksheets
    If Ws.Name = WsIx Then HasWs = True: Exit Function
Next
End Function

Sub Z_WbWcsy()
'D WcStrAyWbOLE(WbzFx(TpFx))
End Sub

Sub Z_LozAyH()
'D NyOy(LozAyH(TpWb))
End Sub

Sub Z_TxtWcCnt()
Dim O As Workbook: 'Set O = WbzFx(Vbe_MthFx)
Ass TxtWcCnt(O) = 1
O.Close
End Sub

Sub Z_SetWcFcsv()
Dim Wb As Workbook
'Set Wb = WbzFx(Vbe_MthFx)
Debug.Print TxtWcStr(Wb)
SetWcFcsv Wb, "C:\ABC.CSV"
Ass TxtWcStr(Wb) = "TEXT;C:\ABC.CSV"
Wb.Close False
Stop
End Sub


Sub Z_Clr_LoRow()
DltLoRow CWs.ListObjects("T_SrcCd")
End Sub
Sub DltLoRow(A As ListObject)
Dim R As Range: Set R = A.DataBodyRange
If IsNothing(R) Then Exit Sub
R.ClearContents
Set R = A1zRg(A.ListColumns(1).Range)
Dim R1 As Range: Set R1 = RgRR(R, 1, 2)
A.Resize R1
End Sub

Sub DltLo(A As Worksheet)
Dim Ay() As ListObject, J%
Ay = IntozItr(Ay, A.ListObjects)
For J = 0 To UB(Ay)
    Ay(J).Delete
Next
End Sub

Sub ClsWsNoSav(A As Worksheet)
ClsWbNoSav WbzWs(A)
End Sub

Sub DltWsIf(A As Workbook, WsIx)
If HasWs(A, WsIx) Then DltWs A, WsIx
End Sub

Sub SavAsCls(Wb As Workbook, Fx)
Wb.SaveAs Fx
Wb.Close
End Sub

Sub SavAsFxm(Wb As Workbook, Fxm)
Wb.SaveAs Fxm, XlFileFormat.xlOpenXMLWorkbookMacroEnabled
End Sub

Function SavAsTmpFxm$(Wb As Workbook)
Dim O$: O = TmpFxm
SavAsFxm Wb, O
SavAsTmpFxm = O
End Function

Function WbnzWs$(A As Worksheet)
WbnzWs = WbzWs(A).FullName
End Function

Sub DltColFm(Ws As Worksheet, FmCol)
WsCC(Ws, FmCol, LasCno(Ws)).Delete
End Sub
Sub DltRowFm(Ws As Worksheet, FmRow)
WsRR(Ws, FmRow, LasRno(Ws)).Delete
End Sub
Sub HidColFm(Ws As Worksheet, FmCol)
WsCC(Ws, FmCol, MaxWsCol).Hidden = True
End Sub

Sub HidRowFm(Ws As Worksheet, FmRow&)
WsRR(Ws, FmRow, MaxWsRow).EntireRow.Hidden = True
End Sub


Function PtCpyToLo(A As PivotTable, At As Range) As ListObject
Dim R1, R2, C1, C2, NC, NR
    R1 = A.RowRange.Row
    C1 = A.RowRange.Column
    R2 = LasRnozRg(A.DataBodyRange)
    C2 = LasCnozRg(A.DataBodyRange)
    NC = C2 - C1 + 1
    NR = R2 - C1 + 1
WsRCRC(WszPt(A), R1, C1, R2, C2).Copy
At.PasteSpecial xlPasteValues

Set PtCpyToLo = CrtLo(RgRCRC(At, 1, 1, NR, NC))
End Function

Sub SetPtffOri(A As PivotTable, FF$, Ori As XlPivotFieldOrientation)
Dim F, J%, T
T = Array(False, False, False, False, False, False, False, False, False, False, False, False)
J = 1
For Each F In Itr(SyzSS(FF))
    With PivFld(A, F)
        .Orientation = Ori
        .Position = J
        If Ori = xlColumnField Or Ori = xlRowField Then
            .Subtotals = T
        End If
    End With
    J = J + 1
Next
End Sub

Sub FmtPt(Pt As PivotTable)

End Sub


Sub ThwHasWbzWs(Wb As Workbook, Wsn$, Fun$)
If HasWs(Wb, Wsn) Then
    Thw Fun, "Wb should have not have Ws", "Wb Ws", Wb.FullName, Wsn
End If
End Sub

Sub SetPtWdt(A As PivotTable, Colss$, ColWdt As Byte)
If ColWdt <= 1 Then Stop
Dim C
For Each C In Itr(SyzSS(Colss))
    EntColzPt(A, C).ColumnWidth = ColWdt
Next
End Sub

Sub SetPtOutLin(A As PivotTable, Colss$, Optional Lvl As Byte = 2)
If Lvl <= 1 Then Stop
Dim F, C As VBComponent
For Each C In Itr(SyzSS(Colss))
    EntColzPt(A, F).OutlineLevel = Lvl
Next
End Sub

Sub SetPtRepeatLbl(A As PivotTable, Rowss$)
Dim F
For Each F In Itr(SyzSS(Rowss))
    PivFld(A, F).RepeatLabels = True
Next
End Sub

Sub ShwPt(A As PivotTable)
ShwXls A.Application
End Sub
Sub PutCd(Cd$(), CdLo As ListObject)
If ChkCdLo(CdLo) Then Exit Sub
DltLoRow CdLo
PutAyV Cd, A1zLo(CdLo)
End Sub
Function ChkCdLo(Lo As ListObject) As Boolean
If IsCdLo(Lo) Then Exit Function
MsgBox "Given Lo is not CdLo", "Lo-Name", Lo.Name
ChkCdLo = True
End Function
Function IsCdLo(A As ListObject) As Boolean
If A.ListColumns.Count <> 1 Then Exit Function
If A.ListColumns(1).Name <> "SrcCd" Then Exit Function
IsCdLo = True
End Function

Function NewA1(Optional Wsn$) As Range
Set NewA1 = A1zWs(NewWs(Wsn))
End Function

Function NewWb(Optional Wsn$) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Add
Set NewWb = WbzWs(SetWsn(FstWs(O), Wsn))
End Function

Function NewWs(Optional Wsn$) As Worksheet
Set NewWs = SetWsn(FstWs(NewWb), Wsn)
End Function

Function NewXls() As Excel.Application
Set NewXls = CreateObject("Excel.Application") ' Don't use New Excel.Application
End Function

Sub Z_NewXls()
Dim X As Excel.Application
Set X = NewXls
Stop
End Sub

Sub QuitXls(A As Excel.Application)
Stamp "QuitXls: Start"
Stamp "QuitXls: ClsAllWb":    ClsAllWb A
Stamp "QuitXls: Quit":        A.Quit
Stamp "QuitXls: Set nothing": Set A = Nothing
Stamp "QuitXls: Done"
End Sub

Sub ClsAllWb(A As Excel.Application)
Dim W As Workbook
For Each W In A.Workbooks
    W.Close False
Next
End Sub


Sub RplLozFb(Wb As Workbook, Fb)
Dim Ws As Worksheet, D As Database
Set D = Db(Fb)
For Each Ws In Wb.Sheets
    RplLozWs Ws, D
Next
D.Close
End Sub

Sub RplLozWs(Ws As Worksheet, D As Database)
Dim Lo As ListObject
For Each Lo In Ws.ListObjects
    RplLozT Lo, D, "@" & Mid(Lo.Name, 3)
Next
End Sub

Sub RplLozT(A As ListObject, D As Database, T)
Dim Fny1$(): Fny1 = Fny(D, T)
Dim Fny2$(): Fny2 = FnyzLo(A)
If Not IsAySam(Fny1, Fny2) Then
    Thw CSub, "LoFny and TblFny are not same", "LoFny T TblFny Db", Fny2, T, Fny1, D.Name
End If
Dim Sq()
    Dim R As DAO.Recordset
    Set R = Rs(A, SqlSel_Fny_T(Fny2, T))
    Sq = AddSngQtezSq(SqzRs(R))
MinxLo A
RgzSq Sq, A.DataBodyRange
End Sub



Sub Z_CrtFxzOupTbl()
Dim Fx$: Fx = TmpFx
CrtFxzOupTbl Fx, SampFbzDutyDta
OpnFx Fx
End Sub

Sub CrtFxzOupTbl(Fx, Fb, Optional AddgWay As EmAddgWay)
SavAsCls NewWbzFbOup(Fb, AddgWay), Fx
End Sub

Function WbzDs(A As Ds) As Workbook
Dim O As Workbook
Set O = NewWb
With FstWs(O)
   .Name = "Ds"
   .Range("A1").Value = A.DsNm
End With
Dim J%, Ay() As Dt
For J = 0 To A.N - 1
    'WszDt O, Ay(J)
Next
Set WbzDs = O
End Function
Sub PutSeqDown(At As Range, FmNum&, ToNum&)
PutAyV LngSeq(FmNum, ToNum), At
End Sub

Sub DltSheet1(Wb As Workbook)
DltWs Wb, "Sheet1"
End Sub
Sub ActWs(Ws As Worksheet)
If IsEqObj(Ws, CWs) Then Exit Sub
Ws.Activate
End Sub
Sub DltWs(Wb As Workbook, WsIx)
Wb.Application.DisplayAlerts = False
If Wb.Sheets.Count = 1 Then Exit Sub
If HasWs(Wb, WsIx) Then WszWb(Wb, WsIx).Delete
End Sub

Sub ClrDown(A As Range)
VbarRgAt(A, AtLeastOneCell:=True).Clear
End Sub


Sub MgeCellAbove(Cell As Range)
'If Not IsEmpty(A.Value) Then Exit Sub
If Cell.MergeCells Then Exit Sub
If Cell.Row = 1 Then Exit Sub
If RgRC(Cell, 0, 1).MergeCells Then Exit Sub
MgeRg RgCRR(Cell, 1, 0, 1)
End Sub


Sub FillSeqH(HBar As Range)
Dim Sq()
Sq = SqVzN(HBar.Rows.Count)
ResiRg(HBar, Sq).Value = Sq
End Sub
Sub ClrCellBelow(Cell As Range)
CellBelow(Cell).Clear
End Sub

Sub FillSeqV(VBar As Range)
Dim Sq()
Sq = SqHzN(VBar.Rows.Count)
ResiRg(VBar, Sq).Value = Sq
End Sub

Sub FillWny(At As Range)
RgzAyV Wny(WbzRg(At)), At
End Sub

Sub FillAtV(At As Range, Ay)
FillSq SqV(Ay), At
End Sub

Sub FillLc(Lo As ListObject, ColNm$, Ay)
If Lo.DataBodyRange.Rows.Count <> Si(Ay) Then Thw CSub, "Lo-NRow <> Si(Ay)", "Lo-NRow ,Si(Ay)", NRowzLo(Lo), Si(Ay)
Dim At As Range, C As ListColumn, R As Range
'DmpAy FnyzLo(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
FillAtV At, Ay
End Sub
Sub FillSq(Sq(), At As Range)
ResiRg(At, Sq).Value = Sq
End Sub
Sub FillAtH(Ay, At As Range)
FillSq SqH(Ay), At
End Sub

Sub RunFxqByCn(Fx, Q)
CnzFx(Fx).Execute Q
End Sub
Function DKVzKSet(KSet As Dictionary) As Drs
Dim K, Dy(): For Each K In KSet.Keys
    Dim S As Aset: Set S = KSet(K)
    Dim V: For Each V In S.Itms
        PushI Dy, Array(K, V)
    Next
Next
DKVzKSet = DrszFF("K V", Dy)
End Function
Sub Z_DKVzLoFilter()
Dim Lo As ListObject: Set Lo = FstLo(CWs)
BrwDrs DKVzLoFilter(Lo)
End Sub
Function DKVzLoFilter(L As ListObject) As Drs
DKVzLoFilter = DKVzKSet(KSetzLoFilter(L))
End Function

Function KSetzKyAsetAy(Ky$(), AsetAy() As Aset) As Dictionary
Set KSetzKyAsetAy = New Dictionary
Dim K, J&: For Each K In Itr(Ky)
    KSetzKyAsetAy.Add K, AsetAy(J)
    J = J + 1
Next
End Function
Sub SetOnFilter(L As ListObject)
On Error GoTo X
Dim M As Boolean: M = L.AutoFilter.FilterMode ' If filter is on, it will have no err, otherwise, there is err
Exit Sub
X:
L.Range.AutoFilter 'Turn on
End Sub
Function KSetzLoFilter(L As ListObject) As Dictionary
'Ret : KSet
Dim O As Dictionary: Set O = New Dictionary
SetOnFilter L
Dim Fny$(): Fny = FnyzLo(L)
Dim F As Filter, J%: For Each F In L.AutoFilter.Filters
    Dim K$: K = Fny(J)
    KSetzLoFilter__Add O, K, F
    J = J + 1
Next
Set KSetzLoFilter = O
End Function

Sub KSetzLoFilter__Add(OKSet As Dictionary, K$, F As Filter)
If Not F.On Then Exit Sub
If F.Operator <> xlFilterValues Then Exit Sub
Dim S As Aset: Set S = AsetzAy(AmRmvPfx(F.Criteria1, "="))
OKSet.Add K, S
End Sub

Function NewWbzDrs(D As Drs) As Workbook
Set NewWbzDrs = WbzRg(RgzDrs(D, NewA1))
End Function

Function RgzDrs(A As Drs, At As Range) As Range
Set RgzDrs = RgzSq(SqzDrs(A), At)
End Function

Function LozDrs(A As Drs, At As Range, Optional Lon$) As ListObject
Set LozDrs = CrtLo(RgzDrs(A, At), Lon)
End Function

Function WszAy(Ay, Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet, R As Range
Set O = NewWs(Wsn)
O.Range("A1").Value = "Array"
Set R = RgzSq(SqV(Ay), O.Range("A2"))
CrtLo RgMoreTop(R)
Set WszAy = O
End Function

Function WszDrs(A As Drs, Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(Wsn)
Dim L As ListObject: Set L = LozDrs(A, O.Range("A1"))
Dim Lc As ListColumn: For Each Lc In L.ListColumns
    StdFmtLc Lc
Next
Set WszDrs = O
End Function

Function RgzAyV(Ay, At As Range) As Range
Set RgzAyV = RgzSq(SqV(Ay), At)
End Function

Function RgzAyH(Ay, At As Range) As Range
Set RgzAyH = RgzSq(SqH(Ay), At)
End Function

Function RgzDy(Dy(), At As Range) As Range
Set RgzDy = RgzSq(SqzDy(Dy), At)
End Function

Function WszDy(Dy(), Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(Wsn)
RgzDy Dy, A1zWs(O)
Set WszDy = O
End Function

Function WszDs(A As Ds) As Worksheet
Dim O As Worksheet: Set O = NewWs
A1zWs(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = WsRC(O, 2, 1)
Dim BelowN&, Dt As Dt, Ay() As Dt
Ay = A.Ay
For J = 0 To A.N - 1
    Dt = Ay(J)
    LozDt Dt, At
    BelowN = 2 + Si(Dt.Dy)
    Set At = CellBelow(At, BelowN)
Next
Set WszDs = O
End Function

Function RgzDt(A As Dt, At As Range, Optional DtIx%)
Dim Pfx$: If DtIx > 0 Then Pfx = QteBkt(CStr(DtIx))
At.Value = Pfx & A.DtNm
RgzSq SqzDrs(DrszDt(A)), CellBelow(At)
End Function

Function LozDt(A As Dt, At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set LozDt = LozDrs(DrszDt(A), R)
RgRC(R, 0, 1).Value = A.DtNm
End Function


Sub Z_WszDs()
ShwWs WszDs(SampDs)
End Sub

Function EnsWbn(Wbn$, InXls As Excel.Application) As Workbook
Dim O As Workbook
Const FxFn$ = "Insp.xlsx"
If HasWbn(InXls, FxFn) Then
    Set O = InXls.Workbooks(FxFn)
Else
    Set O = InXls.Workbooks.Add
    O.SaveAs InstFdr("Insp") & "Insp.xlsx"
End If
Set EnsWbn = ShwWb(O)
End Function

Sub StopXls()
Const Ps1Str$ = "Stop-Process -Id{try{(Get-Process -Name Excel).Id}finally{@()}}.invoke()"
Dim F$: F = TmpHom & "StopXls.ps1"
Static X As Boolean
If Not X Then
    X = True
    If NoFfn(F) Then
        WrtStr Ps1Str, F
    End If
End If
'PowerRun F
End Sub
