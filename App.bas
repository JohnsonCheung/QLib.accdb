VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Const CLib$ = "QApp."
Const CMod$ = CLib & "App."
Private Type A
    H As String
    N As String
    V As String
    OH As String
End Type
Private A As A

Friend Sub Ini(AppHom$, Appn$, Appv$, OupHom$)
A.H = AppHom
A.N = Appn
A.V = Appv
A.OH = OupHom
End Sub

Property Get Appn$()
Appn = A.N
End Property

Function TpFx$(Optional Tpn$)
TpFx = Pth & TpFxFn(Tpn)
End Function

Function TpFxm$(Optional Tpn$)
TpFxm = Hom & TpFxmFn(DftTpn(Tpn))
End Function

Private Function DftTpn$(Tpn$)
DftTpn = IIf(Tpn = "", A.N, Tpn)
End Function

Property Get Pth$()
On Error GoTo E
Static O$
If O = "" Then O = AddFdrAp(A.H, A.N, A.V)
Pth = O
E:
End Property

Function TpFxFn$(Tpn$)
TpFxFn = Tpn & "(Template).xlsx"
End Function

Function TpFxmFn$(Tpn$)
TpFxmFn = Tpn & "(Template).xlsm"
End Function

Property Get OupPth$()
On Error GoTo E
OupPth = AddFdrAp(A.OH, A.N, A.V)
E:
End Property

Property Get OupFxzNxt$()
On Error GoTo E
OupFxzNxt = NxtFfnzAva(OupFx)
E:
End Property

Property Get OupFx$()
On Error GoTo E
OupFx = OupPth & A.N & ".xlsx"
E:
End Property

Property Get Db() As Database
On Error GoTo E
Static A As Database, X As Boolean
If Not X Then
    X = True
    Set A = DbzFb(Fb)
End If
Set Db = A
E:
End Property
Property Get Fb$()
On Error GoTo E
'C:\Users\user\Desktop\MHD\SAPAccessReports\TaxExpCmp\TaxExpCmp\TaxExpCmp.1_3..accdb
Fb = Pth & "AppFb.accdb"
E:
End Property

Property Get Hom$()
On Error GoTo E
Static Y$
If Y = "" Then Y = AddFdrApEns(A.H, A.N, A.V)
Hom = Y
E:
End Property

Function AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
'EnsTblSpec

LnkCcm CurrentDb, CUsr = "User"
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Function

Sub ImpTp()
Dim A As New App
Dim T$: T = A.TpFx
Const CSub$ = CMod & "ImpTp"
If T = "" Then
    Inf CSub, "Tp not exist WPth, no Import", "AppNm Tp WPth", A.Fb, A.Appn, T, Pth
    Exit Sub
End If
Dim D As Database: Set D = A.Db
If IsAttOld(D, "Tp", T) Then ImpAtt D, "Tp", T '<== Import
End Sub

Function TpWsCdNy() As String()
'TpWsCdNy = WszFxwCdNy(TpFx)
End Function

Property Get TpPth$()
On Error GoTo E:
TpPth = EnsPth(TmpHom & "Template\")
E:
End Property

Sub RfhTpWc()
RfhFx TpFx, Fb
End Sub

Function TpWb() As Workbook
Set TpWb = WbzFx(TpFx)
End Function
Function Tp$(Optional Tpn$)

End Function
Sub ExpTp(Optional Tpn$)
ExpAtt Db, "Tp", Tp(Tpn)
End Sub

Friend Sub IniSamp()
Ini "C:\Users\User\Documents\", "App1", "1.1", "C:\Users\User\Documents\App1Oup\"
End Sub

