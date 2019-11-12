Attribute VB_Name = "MxPcws"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPcws."
':Pcat: :Cml #Par-chd-At# ! It is a feature to allow a Drs to be shown in a ws
Sub PutPcwsChd(Target As Range)
Stop
Dim UKeyLo As ListObject: Set UKeyLo = LozHasRg(Target):
If IsNothing(UKeyLo) Then Exit Sub
If Not IsUKeyLo(UKeyLo) Then Exit Sub
Static LasRno&
Dim R&: R = Target.Row
If R = LasRno Then
    Exit Sub
Else
    LasRno = R
End If
Dim Ws As Workbook: Set Ws = WszRg(Target)
Dim Wb As Workbook: Set Wb = WbzWs(Ws)
Dim UKeyLon$: UKeyLon = UKeyLo.Name
Dim SrcLon1$: SrcLon1 = SrcLon(UKeyLon)
Dim ChdLon1$: ChdLon1 = ChdLon(SrcLon1)
Dim SrcLo As ListObject: Set SrcLo = LozWb(Wb, SrcLon1)
Dim ChdLo As ListObject: Set ChdLo = LozWs(Ws, ChdLon1)
'-----------------------------------------------------------------------------------------------------------------------
Dim ChdDy(), KeyDy()
    Dim ChdFny$(): ChdFny = FnyzLo(ChdLo)
    Dim KeyFny$(): KeyFny = FnyzLo(UKeyLo)
    Dim SrcSq(): SrcSq = SqzLo(SrcLo)
    Dim SrcFny$(): SrcFny = FnyzLo(SrcLo)
    Dim ChdCny%(): ChdCny = CnyzSubFny(SrcFny, ChdFny)
    Dim KeyCny%(): KeyCny = CnyzSubFny(SrcFny, KeyFny)
    ChdDy = DyzSqCny(SrcSq, ChdCny)
    KeyDy = DyzSqCny(SrcSq, KeyCny)
    
    Dim KeyDr():        KeyDr = DrzLoCell(UKeyLo, Target)
    Dim CurChdDy():  CurChdDy = DywKeyDr(ChdDy, KeyDr, KeyDy)
    Dim CurChdDrs As Drs:   CurChdDrs = Drs(ChdFny, CurChdDy)
:                              PutDrsToLo CurChdDrs, ChdLo '<===
End Sub

Function IsUKeyLo(L As ListObject) As Boolean
IsUKeyLo = HasSfx(L.Name, "_UKey")
End Function

Function SrcLon$(UKeyLon$)
SrcLon = RmvSfx(UKeyLon, "_UKey")
End Function

Function ChdLon$(SrcLon$)
ChdLon = SrcLon & "_Chd"
End Function
