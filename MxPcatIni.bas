Attribute VB_Name = "MxPcatIni"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPcatIni."
Type Pcat  ' #Par-Chd-At#
    'Key = Ws UKeyRRCC ! Sam Ws, no UKeyRRCC should overlap
    UKeyRRCC As RRCC
    Ws As Worksheet
    '
    ChdFny() As String
    ChdDy() As Variant
    KeyDy() As Variant
    UKeyLo As ListObject
    ChdLo As ListObject
End Type

Sub IniPcat(D As Drs, At As Range, KK$, ShwFF0$, WsSrcl$)
Dim KeyFny$(): KeyFny = AyIntersect(SyzSS(KK), D.Fny): If Si(KeyFny) = 0 Then Thw CSub, "no fld in @KK is fnd in @Drs", "KK D.Fny", KK, D.Fny
Dim ChdFny$(): ChdFny = F_ChdFny1(D.Fny, KeyFny, ShwFF0)

Dim ChdAt As Range:  Set ChdAt = At(1, Si(KeyFny) + 2)
Dim EmpChdDrs As Drs: EmpChdDrs = Drs(ChdFny, EmpAv)
Dim UKeyDrs As Drs: UKeyDrs = SelDistFny(D, KeyFny)
Dim UKeyLo As ListObject: Set UKeyLo = CrtLozDrs(UKeyDrs, At)       '<==
Dim ChdLo As ListObject: Set ChdLo = CrtLozDrs(EmpChdDrs, ChdAt)  '<==
Dim A As Pcat
With A
    Set .Ws = WszRg(At)
    .UKeyRRCC = RRCCzLo(UKeyLo)
    .ChdDy = SelDrsFny(D, ChdFny).Dy
    .ChdFny = ChdFny
    .KeyDy = SelDrsFny(D, KeyFny).Dy
    Set .UKeyLo = UKeyLo
    Set .ChdLo = ChdLo
End With
Push_Pcat_ToPcatBfr A  '<==
AddWsSrc WszRg(At), WsSrcl '<==
AddRfPj PjzRg(At), Pj("QLib") '<==
End Sub

Function F_ChdFny1(Fny$(), KeyFny$(), ShwFF0$) As String()
If ShwFF0 = "" Then
    F_ChdFny1 = AyMinus(Fny, KeyFny)
    Exit Function
End If
F_ChdFny1 = AyIntersect(SyzSS(ShwFF0), Fny)
If Si(F_ChdFny1) Then Thw CSub, "No ShwFld", "Fny KeyFny ShwFF0", Fny, KeyFny, ShwFF0
End Function

Sub Z_IniPcat()
Dim D As Drs, At As Range, KK$, ShwFF0$
GoSub Z
Exit Sub
Z:
    Dim L$: L = Srcl(Md("MxPcatSrc"))
    D = DoMdP
    KK = "CLibv CNsv"
    Set At = NewA1
    IniPcat D, At, KK, ShwFF0, L
    Return
End Sub

