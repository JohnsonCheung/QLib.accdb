Attribute VB_Name = "MxObjPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxObjPrp."

Function PvAyzC(Obj, PrpcAy$()) As Variant()
Dim Prpc: For Each Prpc In PrpcAy
    Push PvAyzC, PvzC(Obj, Prpc)
Next
End Function

Function PvAv(Obj, PrpcAy$()) As Variant()
Const CSub$ = CMod & "DrzObjPrpNy"
If IsNothing(Obj) Then Inf CSub, "Given object is nothing", "PrpcAy", PrpcAy: Exit Function
Dim Prpc: For Each Prpc In PrpcAy
    Push PvAv, Prpc(Obj, Prpc)
Next
End Function

Function DiPrpcqPv(Obj As Object, PrpcAy$()) As Dictionary
Dim DiczObjPrpcAy As New Dictionary
Dim Prpc: For Each Prpc In PrpcAy
    DiczObjPrpcAy.Add Prpc, PvzC(Obj, Prpc)
Next
End Function

Function Pv(Obj, P)
Asg CallByName(Obj, P, VbGet), Pv
End Function

Function PvzC(Obj, Prpc)
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim PrpSeg$(): PrpSeg = Split(Prpc, ".")
Dim O
    Set O = Obj
    Dim U%: U = UB(PrpSeg)
    Dim J%: For J = 0 To U - 1     ' U-1 is to skip the last Pth-Seg
        Set O = PvzC(O, PrpSeg(J)) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next
Asg Pv(O, PrpSeg(U)), PvzC ' Last Prp may be non-object, so must use 'Asg'
Exit Function
X:
Dim E$: E = Err.Description
End Function

