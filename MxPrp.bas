Attribute VB_Name = "MxPrp"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxPrp."

Function DrszItrPrpcc(Itr, Prpcc$) As Drs
Stop
Dim PrpcAy$(): PrpcAy = SyzSS(Prpcc)
    Dim Obj, Dy(): For Each Obj In Itr
        PushI Dy, DrzObjPrpcAy(Obj, PrpcAy)
    Next
DrszItrPrpcc = Drs(PrpcAy, Dy)
End Function

Function DrszItrPrpcAy(Itr, PrpcAy$()) As Drs
DrszItrPrpcAy = Drs(PrpcAy, DyoItrPrpcAy(Itr, PrpcAy))
End Function

Function DrzObjPrpcAy(Obj, PrpcAy$()) As Variant()
Dim P: For Each P In Itr(PrpcAy)
    Push DrzObjPrpcAy, PvzC(Obj, P)
Next
End Function

Function DyoItrPrpcAy(Itr, PrpcAy$()) As Variant()
Dim Obj As Object
For Each Obj In Itr
    Push DyoItrPrpcAy, DrzObjPrpcAy(Obj, PrpcAy)
Next
End Function

Function P_QuietEmp(Obj, P)
On Error Resume Next
Asg CallByName(Obj, P, VbGet), P_QuietEmp
End Function

Function PrpzP1(Obj, P, Optional ThwEr As EmThw)
Select Case True
Case ThwEr = EiNoThw: Asg P_QuietEmp(Obj, P), PrpzP1
Case Else: Stop
End Select
End Function

Sub WAsg3PP(PP_with_NewFldEqQteFmFld$, OPPzPrp$(), OPPzFml$(), OPPzAll$())
Dim I, S$
For Each I In SyzSS(PP_with_NewFldEqQteFmFld)
    S = I
    If HasSubStr(S, "=") Then
        PushI OPPzAll, Bef(S, "=")
        PushI OPPzFml, I
    Else
        PushI OPPzAll, I
        PushI OPPzPrp, I
    End If
Next
End Sub

Function WFmlEr(PrpVy$(), PPzFml$()) As String()
Dim Fml, ErPmAy$(), PmAy$(), O$()
For Each Fml In Itr(PPzFml)
    PmAy = SplitComma(BetBkt(Fml))
    ErPmAy = AyMinus(PmAy, PrpVy)
    If Si(ErPmAy) > 0 Then PushI O, FmtQQ("Invalid-Pm[?] in Fml[?]", JnSpc(ErPmAy), Fml)
Next
If Si(O) > 0 Then PushI O, FmtQQ("Valid-Pm[?]", JnSpc(PrpVy))
WFmlEr = O
End Function


Sub Z_DrszItrPrpcAy()
'BrwDrs DrszItrPrpcc(Excel.Application.Addins, "Name Installed IsOpen FullName CLSId ")
'BrwDrs DrszItrPrpcc(Fds(Db(SampFbzDutyDta), "Permit"), "Name Type Required")
'BrwDrs ItrPrpDrs(Application.VBE.VBProjects, "Name Type")
BrwDrs DrszItrPrpcAy(CPj.VBComponents, SyzSS("Name Type CmpTy=ShpCmpTy(Type)"))
End Sub
