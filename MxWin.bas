Attribute VB_Name = "MxWin"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxWin."


Property Get BrwObjWin() As VBIDE.Window
Set BrwObjWin = FstWin(vbext_wt_Browser)
End Property

Property Get ImmWin() As VBIDE.Window
Set ImmWin = FstWin(vbext_wt_Immediate)
End Property

Property Get LclWin() As VBIDE.Window
Set LclWin = FstWin(vbext_wt_Locals)
End Property

Function WinzMdn(Mdn) As VBIDE.Window
If HasCmpzN(Mdn) Then Set WinzMdn = Md(Mdn).CodePane.Window
End Function

Function WinzM(M As CodeModule) As VBIDE.Window
Set WinzM = M.CodePane.Window
End Function


Sub JmpNxtStmt()
Const CSub$ = CMod & "JmpNxtStmt"
With BoJmpNxtStmt
    If Not .Enabled Then
        'Msg CSub, "BoJmpNxtStmt is disabled"
        Exit Sub
    End If
    .Execute
End With
End Sub

Property Get VisWinCnt&()
VisWinCnt = NItrPrpTrue(CVbe.Windows, "Visible")
End Property
Function CvWin(A) As VBIDE.Window
Set CvWin = A
End Function
Property Get WinCnt&()
WinCnt = CVbe.Windows.Count
End Property

Function MdnCdWin$(CdWin As VBIDE.Window)
MdnCdWin = IsBet(CdWin.Caption, " - ", " (Code)")
End Function

Property Get WinNy() As String()
Dim W As VBIDE.Window
For Each W In CVbe.Windows
    Debug.Print W.Caption, W.Type
    PushI WinNy, W.Caption
Next
End Property

Function FstWin(A As vbext_WindowType) As VBIDE.Window
Dim X As VBIDE.Windows: Set X = CVbe.Windows
Set FstWin = FstzItrEq(CVbe.Windows, "Type", A)
End Function

Function WinyWinTy(T As vbext_WindowType) As VBIDE.Window()
WinyWinTy = IwEq(CVbe.Windows, "Type", T)
End Function

Sub Z_Md()
Debug.Print Md("MVb_Ide_Z_Win").Parent.Name
End Sub

Function PnezCmpn(Cmpn$) As CodePane
Set PnezCmpn = Md(Cmpn).CodePane
End Function

Function WinzCmpn(Cmpn$) As VBIDE.Window
Set WinzCmpn = PnezCmpn(Cmpn).Window
End Function

Sub ClsWinExlMdn(ExlMdn$)
ClsWinExlAp ImmWin, WinzMdn(ExlMdn)
End Sub

Function WinyAv(Av()) As VBIDE.Window()
Dim I
For Each I In Itr(Av)
    PushObj WinyAv, I
Next
End Function

Property Get VisWinCapAy() As String()
VisWinCapAy = SyzOyP(VisWiny, "Caption")
End Property

Property Get VisWiny() As VBIDE.Window()
Dim W As VBIDE.Window: For Each W In CVbe.Windows
    If W.Visible Then PushObj VisWiny, W
Next
End Property

Function CLnozM&(M As CodeModule)
CLnozM = RRCCzPne(M.CodePane).R1
End Function

Function RRCCzPne(P As CodePane) As RRCC
If IsNothing(P) Then Exit Function
Dim R1&, R2&, C1&, C2&
P.GetSelection R1, R2, C1, C2
RRCCzPne = RRCC(R1, R2, C1, C2)
End Function

Function CMthnzM$(M As CodeModule)
Dim K As vbext_ProcKind
CMthnzM = M.ProcOfLine(CLnozM(M), K)
End Function
