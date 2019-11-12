Attribute VB_Name = "MxApplyFilter"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxApplyFilter."
Enum EmCntCol: EiNoCntCol: EiWiCntCol: End Enum
Type RnyPair
    A() As Long
    B() As Long
End Type
Enum EmOp
   EiNop   ' No operation
   EiPatn
   EiEQ    '=
   EiNE    '!
   EiBet   '%
   EiNBET  '!%
   EiLIS   ':
   EiNLIS  '!:
   EiGE    '>=
   EiGT    '>
   EiLE    '<=
   EiLT    '<
End Enum
Const CriCmd$ = "Filter Apply Clear Ins_Row_At Rmv_NRow_At Load"

Function XCriCmdAy() As String()
XCriCmdAy = SyzSS(CriCmd)
End Function

Sub InsCri(L As ListObject)
'Ret : Insert :CriRg: of 3 rows (1 tit and 2 enties) above @L.
'      #1 Ins 3 rows
'      #2 Put lbl
'      #3 Set cmt
'      #4 Add cd [ws on sel chg]
'      #5 Add rf
'      :CriRg: is the rectangle rg just above a list object.  It has 2 or more rows
'      Fst row is tit and the rst are user input area.
'      Fst row cells have value [Select Apply Clear Load]
'              [Select] move the current cell to it, the pgm will tgl to hid/shw the user input area.
'              [Apply]  ..                         , the pgm will apply the cri to the list object.
'                                                    each apply the cri will saved and ok to be retrieve by [Load].
'              [Clear]  ..                         , the pgm will clear the cri value in the range.
'              [Load]   ..                           a new ws show all the applied cri to allow user to select.  After sel
'                                                    the ws will be deleted.
Dim Cell As Range:  Set Cell = A1zRg(L.Range)   ' the A1 cell of @L
Dim W As Worksheet:    Set W = WszLo(L)
Dim M As CodeModule:   Set M = CmpzWs(W).CodeModule ' The md for the ws of @L
Dim Cd$:                  Cd = RplVBar("Sub Worksheet_SelectionChange(ByVal Target As Range)|ApplyFilter Target|End Sub")               ' the code for ws on select
Dim Cmt$:                Cmt = "Explanation of how to use filter"
Dim Pjf$:                Pjf = Pj("QLib").FileName
Dim P As VBProject:    Set P = PjzWs(W)
Ins_Row_At Cell, 3
PutSSH CriCmd, Cell ' Filter Apply Clear Ins_Row_At Rmv_NRow_At Load
SetCmt Cell, Cmt
RplMd M, Cd
P.References.AddFromFile Pjf
End Sub

Sub Z_Apply()
If False Then
    Dim Ws As Worksheet: Set Ws = WsoMthcP
                                  InsCri FstLo(Ws)
    Dim Fx$:                 Fx = SavAsTmpFxm(WbzWs(Ws))
End If
Dim Rg As Range: Set Rg = CWs.Range("B1")
ApplyFilter Rg
End Sub

Function XFCellzL(L As ListObject)
'Ret : :FCell #Filter-Cell fm Lo
Dim A1 As Range: Set A1 = A1zLo(L)
Dim R&: For R = A1.Row - 1 To 1 Step -1
    If RgRC(A1, R, 1).Value = "Filter" Then
        Set XFCellzL = RgRC(A1, R, 1)
        Exit Function
    End If
Next
End Function

Function XLozFCell(FCell As Range) As ListObject
Dim L As ListObjects: Set L = WszRg(FCell).ListObjects
Dim SamCol() As ListObject: SamCol = XSamCol(L, FCell.Column)
If Si(SamCol) = 0 Then Exit Function
#If True Then
Dim R&():         R = XRnyzSamCol(SamCol) ' Rno
Dim MinRno&: MinRno = MinEle(R)
      Set XLozFCell = XLozWhereR(SamCol, MinRno)
#Else
Dim R&():         R = LngAyzOyPrp(SamCol, "Range.Row") ' Rno
Dim MinRno&: MinRno = MinEle(R)
      Set XLozFCell = FstzObj(SamCol, "Range.Row", MinRno)
#End If
End Function

Function XCriRgzL(L As ListObject) As Range
'Ret : :CriRg fm @L ! used by testing
Dim F As Range: Set F = XFCellzL(L)
Set XCriRgzL = XCriRgzFL(F, L)
End Function

Function XCriRgzFL(FCell As Range, L As ListObject) As Range
Dim R2&:   R2 = L.ListColumns(1).Range.Row - 1
Dim C2&:   C2 = L.ListColumns.Count
Set XCriRgzFL = RgRCRC(FCell, 2, 1, R2, C2)
End Function

Sub ApplyFilter(ByVal Tar As Range)
'@ Tar :            ! It is the-cur-active-cell just moved to by user.
'Ret     : do-the-cmd ! cmd comes from the @Tar.value and @Tar comes from %Sub:Worksheet_ChangeSeletion%.
'                     ! If the @Tar is a Filt-cmd-cell, carry the cmd else nop
'-- Fnd L & CriRg ---
Dim FCell As Range: Set FCell = XFCellzT(Tar) '  # Filter-Cell ! the cell with str-"Filter"
                                If IsNothing(FCell) Then Exit Sub
Dim L As ListObject:    Set L = XLozFCell(FCell)
                                If IsNothing(L) Then Debug.Print "ApplyFilter: with FCell, but no Lo": Exit Sub
Dim CriRg As Range: Set CriRg = XCriRgzFL(FCell, L)
'Const  CriCmd$ = "Filter Apply Clear Ins_Row_At Rmv_NRow_At Load"
Select Case Tar.Value ' The Cmd is listed in %CriCmd
Case "Filter": CmdTglFilt CriRg
Case "Apply":  CmdApply L, CriRg: CmdSav CriRg, L
Case "Clear":  CriRg.Clear
Case "Ins_Row_At": Ins_Row_At Tar
Case "Rmv_NRow_At": RmvEmpRow CriRg
Case "Load":   CmdLoad CriRg
Case Else: Debug.Print "ApplyFIlter...should not reach here.'"
End Select
End Sub

Function VzXStr(CellStr)
'Ret : ! a val (Str|Dbl|Bool|Dte|Empty) fm @CellStr.
'      ! having following lgc to determine the val type
'      ! case CellStr
'      ! true or false:           :b
'      ! fstChr='                 :s
'      ! fstChr not (dig and sgn) :s
'      ! can cv to dte:           :dte
'      ! can cv to dbl:           :dbl
'      ! else                     :s
Select Case CellStr
Case "True", "False": VzXStr = CBool(CellStr)
Case Else
    Dim F$: F = FstChr(CellStr)
    If F = "'" Then
        VzXStr = Mid(CellStr, 2)
    ElseIf Not IsAscDigSgn(Asc(F)) Then
        VzXStr = CellStr
    Else
        Dim V: V = CvDbl(CellStr)
        If Not IsEmpty(V) Then
            VzXStr = V
        Else
            V = CvDte(CellStr)
            If Not IsEmpty(V) Then
                VzXStr = V
            Else
                VzXStr = CStr(CellStr)
            End If
        End If
    End If
End Select
End Function


Sub CmdSav(CriRg As Range, L As ListObject)
Exit Sub
Dim Ws As Worksheet: Set Ws = WszRg(CriRg)
Dim S$: S = SqStrzRg(RgMoreTop(CriRg, -1))
Dim Snm$: Snm = CriSnm(L)
Dim D$: D = SzSnm(Ws, Snm)
Dim N$: N = NewCriStr(D, S)
Stop
End Sub
Function NewCriStr$(Saved$, NewStr$)
'Fm Saved  : :CriStr: saved in :Ws-Names:
'Fm NewStr : :CriStr: new cri str to be saved
'Ret : new cri str aft saving @NewStr to @Saved
End Function
Sub ClnUpSavedCri(B As Workbook)
Dim S As Worksheet: For Each S In B.Sheets
    ClnUpSavedCrizS S
Next
End Sub
Function CvWsn$(S As Worksheet)
CvWsn = CvNm(S.Name)
End Function
Function CvLon$(L As ListObject)
CvLon = CvNm(L.Name)
End Function
Function CriSnmPfx$(L As ListObject)
CriSnmPfx = "Cri_" & CvWsn(WszLo(L)) & "_" & CvLon(L) & "_"
End Function
Sub ClnUpSavedCrizS(S As Worksheet)
Dim PfxAy$()
    Dim L As ListObject: For Each L In S.ListObjects
        PushI PfxAy, CriSnmPfx(L)
    Next
RmvSnm S, PfxAy
End Sub
Sub RmvSnm(S As Worksheet, SnmPfxAy$())
Dim N As Excel.Name: For Each N In S.Names
    If HasPfxAy(N.Name, SnmPfxAy) Then N.Delete
Next
End Sub
Function CriSnm$(L As ListObject)
'Ret : :CriSnm: ! :CriSnm: is :Ws-Names-Name: in format of Cri_<NmStr-of-Wsn>_<NmStr-of-Lon>_<TimId-when-it-saved>.
CriSnm = CriSnmPfx(L) & NowId
End Function
Function CvNm$(S$)
'Ret : ret a nm fm @S by rpl non-nm-chr to [_].  Put [N] in front if fst is non-nm-chr
Dim O$: O = S
Dim J%: For J = 1 To Len(O)
    If Not IsNmChr(Mid(O, J, 1)) Then
        Mid(O, J, 1) = "_"
    End If
Next
If Not IsLetter(FstChr(O)) Then
    O = "N" & O
End If
CvNm = O
End Function

Function IxzSprp%(S As Worksheet, Sprpn$)
Dim O%: For O = 1 To S.CustomProperties.Count
    If S.CustomProperties.Item(O).Name = Sprpn Then IxzSprp = O: Exit Function
Next
End Function

Function Sprp(S As Worksheet, Sprpn$)
Dim Ix&: Ix = IxzSprp(S, Sprpn)
If Ix <> 0 Then Sprp = S.CustomProperties.Item(Ix)
End Function
Sub EnsSprp(S As Worksheet, Sprpn$, V)
'Ret : ens a @Sprpnm in @S being val @V.  :Sprpn: is #wS-doc-PRP-Nm.  @@
Dim Ix&: Ix = IxzSprp(S, Sprpn)
If Ix = 0 Then
    S.CustomProperties.Add Sprpn, V
Else
    If S.CustomProperties(Ix).Value <> V Then
        S.CustomProperties(Ix).Value = V
    End If
End If
End Sub

Sub EnsSnm(Ws As Worksheet, Snm$, S$)
'Ret : ens a @Snm in @S being const @S.  @@
If HasSnm(Ws, Snm) Then
    If SzSnm(Ws, Snm) <> S Then Ws.Names(Snm).RefersTo = "=" & QteVb(S)
Else
    Ws.Names.Add Snm, "=" & QteVb(S)
End If
End Sub

Function SzSnm$(S As Worksheet, Snm$)
'Ret : a str stored in ws name stated as @S & @Snm if any.  If not fnd ret blnk & inf.
'      Assume the val stored is a :QteVb:.  If not, ret blnk & inf
If Not HasSnm(S, Snm, CSub) Then Exit Function
Dim O$: O = S.Names(Snm).RefersTo
If Fst2Chr(O) <> "=""" Then Inf CSub, "Ws-Names-Nm exist, but it is not a str cnst", "Ws Snm Stored-RfTo", S.Name, Snm, O: Exit Function
SzSnm = UnQteVb(RmvFstChr(O))
End Function

Function HasSnm(S As Worksheet, Snm$, Optional Fun$) As Boolean
'Fm Fun : is to inf @Snm not fnd if given
'Ret : true if @S has @Snm and inf if not fnd and @Fun is given  @@
':Snm: :Nm #Sheet-Names-Nm#
If HasItn(S.Names, Snm) Then HasSnm = True: Exit Function
If Fun = "" Then Exit Function
Inf Fun, ":Snm:#Ws-Names-Nm is not fnd in ws", "Snm Ws", Snm, S.Name
End Function
'
Sub CmdRmv_NRow_At(CriRg As Range)

End Sub

Sub CmdLoad(CriRg As Range)

End Sub
'
Sub CmdTglFilt(CriRg As Range)

End Sub

Function XFCellzT(Tar As Range) As Range
'@ : :TarCell: is the cell with one of this value: Filter Apply Clear Ins_Row_At Rmv_NRow_At Load, otherwise, don't return :FCell
'Ret : :FCell: is #Filter-Cell ! the cell with str-"Filter" @@
    If Tar.Count <> 1 Then Exit Function
    Dim V:        V = Tar.Value
                      If Not IsStr(V) Then Exit Function
    Dim I%:       I = IxzAy(XCriCmdAy, V)    ' Assume the Cri-Tit-Lin has list of cmd in the order of %CriCmdAy
                      If I = -1 Then Exit Function
    Dim TCno%: TCno = Tar.Column

Dim FCno%: FCno = 1 - I

Dim O As Range: Set O = RgRC(Tar, 1, FCno)
If O.Value <> "Filter" Then Stop
Set XFCellzT = O
'Insp "QXls_Cmd_ApplyFilter.XFCell", "Inspect", "Oup(XFCell) Rg", "NoFmtr(Range)", "NoFmtr(Range)": Stop
End Function

Function XSamCol(A As ListObjects, Cno&) As ListObject()
Dim C As ListObject: For Each C In A
    If C.Range.Column = Cno Then PushObj XSamCol, C
Next
End Function

Function XLozCriRg(CriRg As Range) As ListObject
'Fm CriRg :  # Filter-Cell ! the cell with str-"Filter" @@
If IsNothing(CriRg) Then Exit Function
Dim Ws       As Worksheet:  Set Ws = WszRg(CriRg)
Dim C&:                          C = CriRg.Column
Dim SamCol() As ListObject: SamCol = XSamCol(Ws.ListObjects, C)
Dim R&():                        R = XRnyzSamCol(SamCol)
Dim M&:                          M = MinEle(R)
                     Set XLozCriRg = XLozWhereR(SamCol, M)
End Function

Function XRnyzSamCol(SamCol() As ListObject) As Long()
Dim L: For Each L In Itr(SamCol)
    PushI XRnyzSamCol, CvLo(L).Range.Row
Next
End Function

Function XLozWhereR(A() As ListObject, R&) As ListObject
Dim J%: For J = 0 To UB(A)
    If A(J).Range.Row = R Then Set XLozWhereR = A(J): Exit Function
Next
ThwImpossible CSub
End Function
Sub Z_CmdApply()
Dim Lo As Range, CriRg As Range
GoSub T0
Exit Sub
T0:
    Set Lo = FstLo(CWs)
End Sub

Sub CmdApply(Lo As ListObject, CriRg As Range)
'Fm Lo    : The FilterCell
'Fm CriRg : the rg that user ent to criterior to do filter.
'Ret      : Clr/Set the @CriRg, On & Off vis of row in @Lo.  See Dim Oup* & '<==.

'== ClrMsg==============================================================================================================
DltCmtzRg CriRg     '<==
BdrAroundNone CriRg '<==

'-- Fnd Cri & CriBrk ---------------------------------------------------------------------------------------------------
Dim Fny$():             Fny = FnyzLo(Lo)           ' from Lo
Dim CriCell As Drs: CriCell = XCriCell(Fny, CriRg) ' F R C CriCell ! *CriCell is cri-cell-val

Dim CriBrk As Drs: CriBrk = XCriBrk(CriCell)                                    ' F R C Op V1 V2 Patn IsEr Msg
Dim Cri    As Drs:    Cri = DwEqSel(CriBrk, "IsEr", False, "Op V1 V2 Patn C F") ' Op V1 V2 Patn C F            ! All-IsEr = false

'== Set Cmt & BdrEr to those cri cell with CriStr has err ==============================================================
Dim Er       As Drs:       Er = DwEqSel(CriBrk, "IsEr", True, "R C Msg") ' R C Msg
Dim CellAy() As Range: CellAy = RgAy(CriRg, Er)                          ' Cell with er need to BdrEr and set cmt with msg
Dim Msg$():               Msg = StrCol(Er, "Msg")

Dim OupSetCmt:   SetCmtzAy CellAy, Msg ' <==
Dim OupBdrCriEr: BdrEoAy CellAy       ' <==

'== ShowAllData if there is no Criterior ===============================================================================
If NoReczDrs(Cri) Then
    SetOnFilter Lo
    Dim OupShwAll: Lo.AutoFilter.ShowAllData ' <== if no cri, shw all dta & return
    Exit Sub
End If

'== Apply Filter with Diff =============================================================================================
Dim LRg As Range: Set LRg = Lo.DataBodyRange
    
Dim RFm&:     RFm = LRg.Row                     ' #Fm-Rno-of-the-Lo
Dim RTo&:     RTo = NRowzRg(LRg)                  ' #To-Rno-of-the-Lo
Dim CCny%(): CCny = AwDistAsI(IntCol(Cri, "C")) ' #Cri-Cny          ! #With-Cri-Colno-Ay. The Cno with cri entered
Dim CCol():  CCol = ColAyzLo(Lo, CCny)          ' #Cri-Col          ! #With-Cri-Col.  Ayof-@L-col-wi-cri.  1-#CCol-ele is 1-col.  1-col is ay-of-val-of-a-@L-col.

Dim CCri(): CCri = XCCri(Cri, CCny) ' #Cri-Ay ! 1-%CCriAy-ele is 1-DyOf-[Op V1 V2 Patn C F]
                                    '         ! coming from %Cri, where *C is Cno and *F is Fldn they are for ref
                                    '         ! The *I-th-DyOf-%CCriAy will have sam-*C-and-sam-*F.  The *C will eq to %CCny(*I)
                                    '         ! %CCriAy has sam of ele as %CCny
                                    '         ! %CCriAy is used to sel %CCol in &XVisI
                                                         
Dim VisIxy&(): VisIxy = XVisI(CCol, CCri) ' #Vis-Rxy ! Each ele in %VisI is rix from 0..URow.  All are pointing to row should be visble.
                                          '          ! according to %CCol and %Cri
                                        
Dim Ept() As Boolean: Ept = BoolAybT(RFm, RTo, VisIxy) ' #Ept-Rny-Vis      ! The ix is Rno of @Lo
Dim Act() As Boolean: Act = VisAyzLo(Lo)               ' #Act-Rny-Vis      ! The ix is Rno of @Lo
Dim Vis   As RnyPair: Vis = XVis(Ept, Act)             ' #Vis-Hid-Rny-Pair ! %Vis:RnyPair.A is Vis and .B is Hid

Dim Ws As Worksheet: Set Ws = WszLo(Lo)
Dim OupVisHid:                SetVis Ws, Vis ' ! <== Turn on or off the row

'== Bdr the cri selecting no record (Ns) (no-sel) ======================================================================
'Dim NsFny$(): NsFny = XNsFny(CSel, CFny) '  ! What CFny selecting no rec
'If Si(NsFny) > 0 Then
'    Dim NsCri    As Drs:     NsCri = DwIn(CriBrk, "F", NsFny) ' Cri causing the no sel
'    Dim NsExlEr  As Drs:   NsExlEr = DwEq(NsCri, "IsEr", False) ' Xls those already @IsEr
'    Dim NsRpt    As Drs:     NsRpt = SelDrs(NsExlEr, "R C")     ' Need to report ns cri
'    Dim NsCell() As Range:  NsCell = RgAy(CriRg, NsRpt)
'    Dim OupBdrNoRec:                 BdrEoAy NsCell, "Red"      ' <== Bdr the Cri cell which does not sel and record
'End If
End Sub
Function XVisI(CCol(), CCri()) As Long()
'Fm CCol   :  #Cri-Col ! #With-Cri-Col.  Ayof-@L-col-wi-cri.  1-#CCol-ele is 1-col.  1-col is ay-of-val-of-a-@L-col.
'Fm CCriAy :  #Cri-Ay  ! 1-%CCriAy-ele is 1-DyOf-[Op V1 V2 Patn C F]
'                      ! coming from %Cri, where *C is Cno and *F is Fldn they are for ref
'                      ! The *I-th-DyOf-%CCriAy will have sam-*C-and-sam-*F.  The *C will eq to %CCny(*I)
'                      ! %CCriAy has sam of ele as %CCny
'                      ! %CCriAy is used to sel %CCol in &XVisI
'Ret       :  #Vis-Ixy ! Each ele in %VisI is rix from 0..URow.  All are pointing to row should be visble.
'                      ! according to %CCol and %Cri @@
If Si(CCol) <> Si(CCri) Then Stop
Dim CCol0(): CCol0 = CCol(0) ' #Cri-Col-0 ! the fst col of %CCol #Cri-Col
Dim URow&: URow = UB(CCol0)   '            ! same as @L-URow
Dim O&(): O = LngSeq(0, URow) ' ! assume all row are selected
Dim ICol, J%: For Each ICol In CCol
    Dim Col(): Col = ICol
    Dim CriAy(): CriAy = CCri(J)
    O = XRxywCri(O, Col, CriAy)
    J = J + 1
Next
XVisI = O
'Insp "QXls_Cmd_ApplyFilter.XVisI", "Inspect", "Oup(XVisI) CCol CCriAy", XVisI, CCol, CCriAy: Stop
End Function

Function XIsSelAy(SubCol(), CriAy()) As Boolean()
'Fm SubCol : to be further selected by @Cri
'Fm Cri    : Dy-Op-V1-V2-Patn-F-C, where *F is Fldn & *C is Cno for rf.
'Ret : ! %IsSelAy which has sam # of ele as @SubCol.  This fun is further determine they are selected.
Dim ICri: For Each ICri In CriAy
    Dim Op As EmOp:       Op = ICri(0)
    Dim V1:               V1 = ICri(1)
    Dim V2:               V2 = ICri(2)
    Dim Patn$:          Patn = ICri(3)
    Dim Re As RegExp: Set Re = Rx(Patn)
    Dim V: For Each V In SubCol
        Dim IsSel As Boolean: IsSel = XIsSel(V, Op, V1, V2, Re)
        PushI XIsSelAy, IsSel
    Next
Next
End Function
Function XRxywAy(Rxy&(), IsSelAy() As Boolean) As Long()
'-- Use %IsSelAy to select subset of %RxySel as %Ret:XRxywCri.  It will be less or eq @RxySel.
Dim J&: For J = 0 To UB(IsSelAy)
    If IsSelAy(J) Then
        PushI XRxywAy, Rxy(J)
    End If
Next
End Function

Function XRxywCri(RxySel() As Long, Col(), CriAy()) As Long()
'Fm RxySel : #Rxy-Selected          ! it is AyOf-Ix-pointing to %Col.  Each ele is bet 0..UB(%Col).  The ele are pointing selected row.
'                                   ! They need further select by %Cri.
'Fm Col    : #Col-Vy                ! It is cur-col-vy.  It has sam nbr of ele as @L-Rows.
'                                   ! Only sub-set of this %Col is used to be test to selected, which is indicated by @RxySel
'Fm Cri    : :Dy:Op-V1-V2-Patn-C-F !
'Ret       : #Rxy-Further-Selected  ! It is sub set of @RxySel
Dim SubCol():               SubCol = AwIxy(Col, RxySel)     ' ! sam ele as @RxySel
Dim IsSelAy() As Boolean:  IsSelAy = XIsSelAy(SubCol, CriAy)
                          XRxywCri = XRxywAy(RxySel, IsSelAy)
End Function

Function VisAyzLo(L As ListObject) As Boolean()
'Ret : #Vis-Array-Fm-Lo# ! ret a boolean array with true means visible and false means hidden of each row stated in @R.
'                        ! the ix of the @Ret is row no.
VisAyzLo = VisAyzRg(RgzLc(L, 1))
End Function

Sub Z_DrszFTnbr()
Dim R&(): R = SampRny
XBox "SampRny"
X R
XLin
XBox "DrszFTnbr"
X FmtDrszNoRdu(DrszFTnbr(R))
Brw XX
End Sub

Function DrszFTnbr(SrtdNumAy, Optional CntCol As EmCntCol) As Drs
Stop
Dim A$:       If CntCol = EiWiCntCol Then A = " Cnt"
Dim FF$: FF = "Fm To" & A
DrszFTnbr = DrszFF("Fm To", DyoFTnbr(SrtdNumAy, CntCol))
End Function

Function DyoFTnbr(NumAy, Optional CntCol As EmCntCol) As Variant()
'Fm NumAy : assume srted.
'Ret      : :DyoFTnbr: is dry with 2 nbr col fst is *FmNum and snd is *ToNum.
If Si(NumAy) = 0 Then Exit Function
Dim O()
Dim LasF&: LasF = NumAy(0)
Dim LasT&: LasT = LasF
Dim N: For Each N In NumAy
    If N - 1 = LasT Then
        LasT = N
    ElseIf N = LasT Then
    ElseIf N > LasT Then
        PushI O, Array(LasF, LasT)
        LasF = N
        LasT = N
    Else
         Thw CSub, "NumAy is not sorted", "Not-srt-ele NumAy", N, NumAy
    End If
Next
PushI O, Array(LasF, LasT)
If CntCol = EiWiCntCol Then
    Dim J&: For J = 0 To UB(O)
        O(J) = AddEle(O, O(J)(1) - O(J)(0))
    Next
End If
DyoFTnbr = O
End Function

Function SampRny() As Long()
':Rny: :Long() #RowNo-Ay# ! RowNo is started from 1
SampRny = IntozAy(EmpLngAy, Res("SampRny"))
End Function

Function ColontAy(L&()) As String()
ColontAy = ColontAyzFT(DyoFTnbr(L))
End Function
Function ColontAyzFT(DyoFTnbr()) As String()
'Ret : :Colont: is #Colon-Term#
Dim FTnbr: For Each FTnbr In Itr(DyoFTnbr)
    PushI ColontAyzFT, FTnbr(0) & ":" & FTnbr(1)
Next
End Function
Function EntRzWsRny(Ws As Worksheet, Rny&()) As Range
Dim S$: S = JnComma(ColontAy(Rny))
S = "88:90,92:271"
Set EntRzWsRny = Ws.Range(S).EntireRow
End Function

Sub SetViszWsRny(Ws As Worksheet, Rny&(), Vis As Boolean)
If Si(Rny) = 0 Then Exit Sub
#If False Then
Dim R As Range
Dim Rno: For Each Rno In Itr(Rny)
    Set R = Ws.Rows(Rno)
    R.EntireRow.Hidden = Not Vis
Next
#Else
EntRzWsRny(Ws, Rny).Hidden = Not Vis
#End If
End Sub

Sub Z_BrwNumAy()
BrwNumAy SampRny, EiWiCntCol
End Sub

Sub BrwNumAy(SrtdNumAy, Optional CntCol As EmCntCol)
BrwAy FmtNumAy(SrtdNumAy, CntCol)
End Sub

Function FmtNumAy(SrtdNumAy, Optional CntCol As EmCntCol = EiWiCntCol) As String()
FmtNumAy = FmtDrszNoRdu(DrszFTnbr(SrtdNumAy, CntCol))
End Function
Function FmtFnyPair(A As RnyPair, Optional Nm$) As String()
Dim B As New Bfr
B.Box Nm & ":RnyPair"
B.ULin "A-Rny"
B.Var FmtNumAy(A.A)
B.ULin "B-Rny"
B.Var FmtNumAy(A.B)
FmtFnyPair = B.Ly
End Function
Sub SetVis(Ws As Worksheet, VisHid As RnyPair)
Stop
SetViszWsRny Ws, VisHid.A, True
SetViszWsRny Ws, VisHid.B, False
End Sub

Function VisAyzRg(R As Range) As Boolean()
'Ret : #Vis-Array-Fm-Rg# ! ret a boolean array with true means visible and false means hidden of each row stated in @R.
'                        ! the ix of the @Ret is row no.
Dim RFm&: RFm = R.Row
Dim RTo&: RTo = NRowzRg(R)
Dim O() As Boolean: ReDim O(RFm To RTo)
Dim Rg As Range: For Each Rg In RgC(R, 1)
    With Rg
        O(.Row) = Not .EntireRow.Hidden
    End With
Next
VisAyzRg = O
End Function

Function XVis(Ept() As Boolean, Act() As Boolean) As RnyPair
'Fm Ept :  #Ept-Rny-Vis      ! The ix is Rno of @Lo
'Fm Act :  #Act-Rny-Vis      ! The ix is Rno of @Lo
'Ret    :  #Vis-Hid-Rny-Pair ! %Vis:RnyPair.A is Vis and .B is Hid @@
'Fm Ept : #Visible-Array#
'Fm Act : #Visible-Array#
'Ret    : @Ret.A is visible, @Ret.B is hidden
Dim Vis&(), Hid&() ' :Rno:
Dim R&: For R = LBound(Ept) To UBound(Ept)
    Dim E&: E = Ept(R)
    Dim A&: A = Act(R)
    Select Case True
    Case E And Not A: PushI Vis, R
    Case Not E And A: PushI Hid, R
    End Select
Next
XVis.A = Vis
XVis.B = Hid
'Insp "QXls_Cmd_ApplyFilter.XVis", "Inspect", "Oup(XVis) Ept Act", "NoFmtr(RnyPair)", "NoFmtr(() As Boolean)", "NoFmtr(() As Boolean)": Stop
End Function

Function XNsFny(CSel() As Aset, CFny$()) As String()
'Fm CSel :  # Cri-Sel ! aft the Cri apply, what val in col-vset has been selected
'                     ! [CFny CCol CSet] hav sam nbr of ele
'                     ! ele in @CSel may have no ele, it will consider as warning, need to rpt by
'                     ! by bdr er all the cri cell.
'Fm CFny :  # Cri-Fny ! The Fny with cri entered
'Ret     :            ! What CFny selecting no rec @@
Dim I, J%: For Each I In CSel
    If CvAset(I).Cnt = 0 Then
        PushI XNsFny, CFny(J)
    End If
Next
'Insp "QXls_Cmd_ApplyFilter.XNsFny", "Inspect", "Oup(XNsFny) CSel CFny", XNsFny, "NoFmtr(() As Aset)", CFny: Stop
End Function

Sub SetCmtzAy(Rg() As Range, Cmt$())
Dim J%: For J = 0 To MinUB(Rg, Cmt)
    SetCmt Rg(J), Cmt(J)
Next
End Sub

Sub Z_SetCmt()
Dim R As Range: Set R = A1zWs(CWs)
SetCmt R, "lskdfjsdlfk"
End Sub

Function HasCmt(R As Range) As Boolean
HasCmt = Not IsNothing(R.Comment)
End Function

Sub SetCmt(R As Range, Cmt$)
If Not HasCmt(R) Then
    R.AddComment.text Cmt
    Exit Sub
End If
Dim C As Comment: Set C = R.Comment
If C.text = Cmt Then Exit Sub
C.text Cmt
End Sub

Sub Z_BdrEr()
BdrEr CWs.Range("C2"), "Red"
End Sub


Function XCriBrk(CriCell As Drs) As Drs
'Fm CriCell : F R C CriCell                ! *CriCell is cri-cell-val
'Ret        : F R C Op V1 V2 Patn IsEr Msg @@
Dim Dr, Dy(): For Each Dr In Itr(CriCell.Dy)
    Dim CriCellVal: CriCellVal = Pop(Dr)          ' The cell value with criteria val
    Dim CriDr:   CriDr = XCriDrzCellVal(CriCellVal) ' Op V1 V2 Patn IsEr Msg        ! brk the criteria str in these 6 values
    If Si(CriDr) <> 6 Then Stop
    PushIAy Dr, CriDr
    PushI Dy, Dr
Next
XCriBrk = DrszFF("F R C Op V1 V2 Patn IsEr Msg", Dy)
'Insp "QXls_Cmd_ApplyFilter.XCriBrk", "Inspect", "Oup(XCriBrk) CriCell", FmtDrszNoRdu(XCriBrk), FmtDrszNoRdu(CriCell): Stop
End Function

Function XIsSel(V, Op As EmOp, V1, V2, Re As RegExp) As Boolean
'Ret : If @V is selected by @*
On Error GoTo X
Dim O As Boolean
Select Case True
Case Op = EiPatn: O = Re.Test(V)
Case Op = EiBet:  O = IsBet(V, V1, V2)
Case Op = EiNBET: O = Not IsBet(V, V1, V2)
Case Op = EiGE:   O = V >= V1
Case Op = EiGT:   O = V > V1
Case Op = EiLE:   O = V <= V1
Case Op = EiLT:   O = V < V1
Case Op = EiLIS:  O = HasEle(V1, V)
Case Op = EiNLIS: O = Not HasEle(V1, V)
Case Op = EiNE:   O = V <> V1
Case Op = EiEQ:   O = V = V1
Case Else: Thw CSub, "Op error"
End Select
XIsSel = O
Exit Function
X: Dim E$: E = Err.Description
   Inf CSub, "Runtime er", "Er V-To-Be-Select Op V1 V2 Patn", E, V, Op, V1, V2, Re.Pattern
End Function

Function XCriCell(Fny$(), CriRg As Range) As Drs
'Fm Fny : from Lo
'Ret    : F R C CriCell ! *CriCell is cri-cell-val @@
Dim Sq(): Sq = CriRg.Value
Dim CriCol()
Dim C%, F, Dy(): For Each F In Fny
    C = C + 1
    CriCol = ColzSq(Sq, C)
    Dim R%: R = 0
    Dim CriCell: For Each CriCell In CriCol
        R = R + 1
        If Not IsEmpty(CriCell) Then
            PushI Dy, Array(F, R, C, CriCell)
        End If
    Next
Next
XCriCell = DrszFF("F R C CriCell", Dy)
'Insp "QXls_Cmd_ApplyFilter.XCriCell", "Inspect", "Oup(XCriCell) Fny CriRg", FmtDrszNoRdu(XCriCell), Fny, "NoFmtr(Range)": Stop
End Function

Function XCriDr(Op As EmOp, V1, Optional V2, Optional Patn$, Optional IsEr As Boolean, Optional Msg$) As Variant()
XCriDr = Array(Op, V1, V2, Patn, IsEr, Msg)
End Function

Function XCriDrzCellVal(CriCellVal)
'Fm CriVal : The cell value with criteria val
'Ret : DrOf Op-V1-V2-Patn-IsEr-Msg
Dim V: V = CriCellVal
Dim T As EmSimTy: T = SimTy(V)
Dim O()
Select Case T
Case EiYes, EiDte, EiNum: O = XCriDr(EiEQ, V)
Case EiStr:               O = XCriDrzStr(CStr(V))
End Select
XCriDrzCellVal = O
'Insp "QXls_Cmd_ApplyFilter.XCriDr", "Inspect", "Oup(XCriDr) CriVal", XCriDr, CriVal: Stop
End Function

Function XShfOp(OStr$) As EmOp
'   EiPatn
'   EiEQ    '=
'   EiNE    '!
'   EiBET   '%
'   EiNBET  '!%
'   EiLIS   ':
'   EiNLIS  '!:
'   EiGE    '>=
'   EiGT    '>
'   EiLE    '<=
'   EiLT    '<
Dim F$: F = FstChr(OStr)
Dim O As EmOp
Select Case F
Case "=": O = EiEQ: GoTo X1
Case "!"
    If SndChr(OStr) = ":" Then
        O = EiNLIS
        GoTo X2
    End If
    O = EiNE
    GoTo X1
Case "%": O = EiBet: GoTo X1
Case ":": O = EiLIS: GoTo X1
Case ">"
    If SndChr(OStr) = "=" Then
        O = EiGE
        GoTo X2
    End If
    O = EiGT
    GoTo X1
Case "<"
    If SndChr(OStr) = "=" Then
        O = EiLE
        GoTo X2
    End If
    O = EiLT
    GoTo X1
Case Else
    XShfOp = EiPatn
End Select
Exit Function
X1:
    XShfOp = O: OStr = RmvFstChr(OStr)
    Exit Function
X2:
    XShfOp = O: OStr = Mid(OStr, 3)
End Function

Function XCriDrzBet(Op As EmOp, IsStr As Boolean, S$) As Variant()
Dim V1$, V2$: AsgTRst S, V1, V2
Dim O()
Select Case True
Case IsStr:                       O = XCriDr(Op, V1, V2)
Case IsNumeric(V1) And IsNumeric(V2): O = XCriDr(Op, CDbl(V1), CDbl(V2))
Case IsDtezS(V1) And IsNumeric(V2): O = XCriDr(Op, CDate(V1), CDate(V2))
Case Else:                        O = XCriDr(Op, V1, V2)
End Select
If V1 > V2 Then Inf CSub, "V1 cannot > V2", "V1-TyNm V1 V2", TypeName(V1), V1, V2: Exit Function
XCriDrzBet = O
End Function

Function XCriDrzLis(Op As EmOp, IsStr As Boolean, S$) As Variant()
Dim Sy$(): Sy = SyzSS(S)
If IsSyDte(Sy) Then XCriDrzLis = XCriDr(Op, DteAyzSy(Sy)): Exit Function
If IsSyDbl(Sy) Then XCriDrzLis = XCriDr(Op, DblAyzSy(Sy)): Exit Function
XCriDrzLis = XCriDr(Op, Sy)
End Function

Function XCriDrzSng(Op As EmOp, IsStr As Boolean, S$) As Variant()
Dim A$: A = Trim(S)
If IsStr Then
    XCriDrzSng = XCriDr(Op, S)
    Exit Function
End If
If IsDtezS(S) Then
    Dim D As Date: D = A
    XCriDrzSng = XCriDr(Op, D)
    Exit Function
End If

If IsNumeric(S) Then
    Dim Dbl#: Dbl = A
    XCriDrzSng = XCriDr(Op, Dbl)
    Exit Function
End If
XCriDrzSng = XCriDr(Op, A)
End Function
Function XCriDrzStr(S$) As Variant()
'Ret : DrOf Op V1 V2 Patn IsEr Msg
Dim Op As EmOp: Op = XShfOp(S)
If S = "" Then
    Inf CSub, "S should be blank"
    Exit Function
End If
Dim L$: L = S
Dim IsStr As Boolean: If Op <> EiPatn Then If ShfPfx(L, "'") Then IsStr = True
L = Trim(L)
Select Case Op
Case EiBet, EiNBET:                      XCriDrzStr = XCriDrzBet(Op, IsStr, L)
Case EiLE, EiLT, EiGT, EiGE, EiNE, EiEQ: XCriDrzStr = XCriDrzSng(Op, IsStr, L)
Case EiLIS, EiNLIS:                      XCriDrzStr = XCriDrzLis(Op, IsStr, L)
Case EiPatn:                             XCriDrzStr = XCriDr(Op, Empty, Patn:=L)
Case Else:                               XCriDrzStr = XCriDrzEr("Invalid EmOp from CellVal")
End Select
End Function
Function XCriDrzEr(Msg$) As Variant()
XCriDrzEr = XCriDr(EiNop, "", IsEr:=True, Msg:=Msg)
End Function

Function NotUse_XAsk$()
Dim T$: T = "Filter"
Erase XX
X "[Yes] = Apply"
X "[No] = Clear"
Dim M$: M = JnCrLf(XX)
Dim S As VbMsgBoxStyle: S = vbYesNoCancel + vbQuestion + vbDefaultButton1
Dim R As VbMsgBoxResult: R = MsgBox(M, S, T)
Dim O$
Select Case True
Case R = vbYes: O = "Apply"
Case R = vbNo: O = "Clear"
End Select
NotUse_XAsk = O
End Function

Function XCCri(Cri As Drs, CCny%())
'Fm Cri  : Op V1 V2 Patn C F          ! All-IsEr = false
'Fm CCny :                   #Cri-Cny ! #With-Cri-Colno-Ay. The Cno with cri entered
'Ret     :                   #Cri-Ay  ! 1-%CCriAy-ele is 1-DyOf-[Op V1 V2 Patn C F]
'                                     ! coming from %Cri, where *C is Cno and *F is Fldn they are for ref
'                                     ! The *I-th-DyOf-%CCriAy will have sam-*C-and-sam-*F.  The *C will eq to %CCny(*I)
'                                     ! %CCriAy has sam of ele as %CCny
'                                     ! %CCriAy is used to sel %CCol in &XVisI @@
'Insp "QXls_Cmd_ApplyFilter.XCCriAy", "Inspect", "Oup(XCCriAy) Cri CCny", XCCriAy, FmtDrszNoRdu(Cri), CCny: Stop
Dim O()
    Dim C: For Each C In CCny
        PushI O, DwEq(Cri, "C", C).Dy
    Next
XCCri = O
'Insp "QXls_Cmd_ApplyFilter.XCCri", "Inspect", "Oup(XCCri) Cri CCny", XCCri, FmtDrszNoRdu(Cri), CCny: Stop
End Function
