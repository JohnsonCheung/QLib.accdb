Attribute VB_Name = "MxPcwsIni"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPcwsIni."

Sub IniPcws(Lo As ListObject, KK$, Adr$, WsSrcl$)
'Do 1:NewWs 2:PutUKey 3:AddEmpChdDrs 4:AddRf 5:AddWsSrc 6:PutChd
Dim LoFny$():   LoFny = FnyzLo(Lo)

S1:
Dim Ws As Worksheet: Set Ws = AddWs(WbzLo(Lo))                  '<== 1-NewWs
S2:
    Dim KeyFny$(): KeyFny = AyIntersect(SyzSS(KK), LoFny)
    Dim At As Range:     Set At = Ws.Range(Adr)
    Dim UKeyDrs As Drs: UKeyDrs = SelDistAllCol(DrszLoFny(Lo, KeyFny))
                            LozDrs UKeyDrs, At, Lo.Name & "_UKey"                  '<== 2:PutUKey
S3: Dim EmpChdDrs As Drs
    Dim ChdFny$(): ChdFny = AyMinus(LoFny, KeyFny)
    Dim ChdAt As Range: Set ChdAt = RgRC(At, 1, Si(KeyFny) + 2)
                   EmpChdDrs = Drs(ChdFny, EmpAv)
                   LozDrs EmpChdDrs, ChdAt, Lo.Name & "_Chd"
S4: AddRfPj PjzWs(Ws), Pj("QLib")
S5: AddWsSrc Ws, WsSrcl '<==5:AddWsSrc

S6:  Dim Tar As Range: Set Tar = RgRC(At, 2, 1)
Stop
    PutPcwsChd Tar

End Sub

Function F_ChdFny(Fny$(), KeyFny$(), ShwFF0$) As String()
If ShwFF0 = "" Then
    F_ChdFny = AyMinus(Fny, KeyFny)
    Exit Function
End If
F_ChdFny = AyIntersect(SyzSS(ShwFF0), Fny)
If Si(F_ChdFny) Then Thw CSub, "No ShwFld", "Fny KeyFny ShwFF0", Fny, KeyFny, ShwFF0
End Function

Sub Z_IniPcws()
Dim At As Range, Lo As ListObject
GoSub Z
Exit Sub
Z:
    ClsAllWbNoSav
    Dim L$: L = Srcl(Md("MxPcwsSrc"))
    Set Lo = ResLo("IniPcws")
    IniPcws Lo, "Pjn CLibv CNsv", "A1", L
    Return
End Sub
