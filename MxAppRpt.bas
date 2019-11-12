Attribute VB_Name = "MxAppRpt"
Option Compare Text
Option Explicit
Const CNs$ = "dfd"
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxAppRpt."
Private Type CWs: FmWsn As String: End Type
Private Type CFx: Fx As String: CWs() As CWs: End Type
Type RptPm
    InpFilSrc() As String
    LnkImpSrc() As String
    WPth As String
    WFb As String
    TpFx As String
    OupFx As String
    OupGenr As IOupGenr
    WbFmtr As IWbFmtr
    IsOpn As Boolean
    IsCpyInp As Boolean
End Type

Function CFxSi%(A() As CFx)
On Error Resume Next
CFxSi = UBound(A) + 1
End Function

Function CFxUB%(A() As CFx)
CFxUB = CFxSi(A) - 1
End Function

Function CWsSi%(A() As CWs)
On Error Resume Next
CWsSi = UBound(A) + 1
End Function

Function CWsUB%(A() As CWs)
CWsUB = CWsSi(A) - 1
End Function

Sub PushCFx(O() As CFx, M As CFx)
Dim N%: N = CFxSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Sub Rpt(P As RptPm) 'Gen&Vis OupFx using LidPm as NxtFfnzAva.
With P
:                              CpyFfnAyzIfDif .InpFilSrc, .WPth         ' <== Cpy inp fil to wpth
:                              EnsFb .WFb                              ' <== Crt wrk fb
Dim W As Database:     Set W = Db(.WFb)
:                              LnkImp Sy(.InpFilSrc, .LnkImpSrc), W    ' <== LnkImp
:                              .OupGenr.GenOupTbl W            ' <== Gen oup tbl
:                              W.Close
:                              CpyFfn .TpFx, .OupFx                    ' <== Cpy to OupFx.  Assume OupFx is always new
:                              RfhFx .OupFx, .WFb
:                              If Not XShouldOpnWb(P) Then Exit Sub    ' <== is it done?
Dim OWb As Workbook: Set OWb = WbzFx(.OupFx)
:                              XCpyInp P, OWb                          ' <== Cpy inp ws
:                              XFmt .WbFmtr, OWb                       ' <== Fmt wb
:                              OWb.Save                                ' <== Sav
:                              If Not .IsOpn Then OWb.Close            ' <== KeepOpn?
End With
End Sub

Sub XCFx(P As CFx, ToWb As Workbook)
Dim FmWb As Workbook: Set FmWb = Xls.Workbooks.Open(P.Fx)
Dim J%: For J = 0 To CWsUB(P.CWs)
    XCpyWs P.CWs(J), FmWb, ToWb
Next
FmWb.Close
End Sub

Function XCFxAy(P As RptPm) As CFx()

End Function

Sub XCpyInp(P As RptPm, ToWb As Workbook)
Dim CFx() As CFx: CFx = XCFxAy(P)
Dim J%: For J = 0 To CFxUB(CFx)
    XCFx CFx(J), ToWb
Next
End Sub

Sub XCpyWs(P As CWs, Fm As Workbook, Tar As Workbook)
Dim FmWs As Worksheet
End Sub

Sub XFmt(F As IWbFmtr, B As Workbook)
If Not IsNothing(F) Then F.FmtWb B
End Sub

Function XShouldOpnWb(P As RptPm) As Boolean
XShouldOpnWb = True
With P
    If Not IsNothing(.WbFmtr) Then Exit Function
    If .IsOpn Then Exit Function
    If .IsCpyInp Then Exit Function
End With
XShouldOpnWb = False
End Function

Sub Z_Rpt()
Dim WDb As Database, B As IOupGenr
GoSub Z
Exit Sub
Z:
    Set WDb = Nothing
    Set B = New TaxCmpOupGenr
    GoTo Tst
Tst:
    'Rpt_Oup WDb, B
    Return
End Sub
