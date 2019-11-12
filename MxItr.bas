Attribute VB_Name = "MxItr"
Option Compare Text
Option Explicit
Const CNs$ = "ObjPrp"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxItr."

Function AvzItr(Itr) As Variant()
AvzItr = IntozItr(Array(), Itr)
End Function


Function ItrClnAy(Itr)
If Itr.Count = 0 Then Exit Function
Dim X
For Each X In Itr
    ItrClnAy = Array(X)
    Exit Function
Next
End Function

Function NItrPrpTrue(Itr, BoolPrpNm)
Dim O&, X
For Each X In Itr
    If CallByName(X, BoolPrpNm, VbGet) Then
        O = O + 1
    End If
Next
NItrPrpTrue = O
End Function

Sub ForItrFun(Itr, DoFun$)
Dim I: For Each I In Itr
    Run DoFun, I
Next
End Sub

Sub ForItrFunPX(Itr, PX$, P)
Dim X
For Each X In Itr
    Run PX, P, X
Next
End Sub

Sub ForItrFunXP(Itr, Xp$, P)
Dim X
For Each X In Itr
    Run Xp, X, P
Next
End Sub

Function FstItm(Itr)
Dim X
For Each X In Itr
    Asg X, FstItm
    Exit Function
Next
End Function

Function FstItmPredXP(Ay, Xp$, P$)
Dim X
For Each X In Ay
    If Run(Xp, X, P) Then Asg X, FstItmPredXP: Exit Function
Next
End Function

Function FstzItrEq(Itr, Prpc, V)
'Ret : fst ele in @Itr with its prpOf-@Prpc eq to @V
Dim Obj: For Each Obj In Itr
    If PvzC(Obj, Prpc) = V Then Set FstzItrEq = Obj: Exit Function
Next
Set FstzItrEq = Nothing
End Function

Function FstzItn(Itr, Nm$) 'Return first element in Itr with its PrpNm=Nm being true
Set FstzItn = FstzItrEq(Itr, "Name", Nm)
End Function

Function FstzItrT(Itr, TruePrpc$)
'Ret : fst ele in @Itr wi its prp-of-@TruePrpc being true
Set FstzItrT = FstzItrEq(Itr, TruePrpc, True)
End Function

Function HasItn(Itr, Nm) As Boolean
Dim Obj: For Each Obj In Itr
    If Obj.Name = Nm Then HasItn = True: Exit Function
Next
End Function

Function HasItrEq(Itr, Prpc, V) As Boolean
Dim Obj: For Each Obj In Itr
    If PvzC(Obj, Prpc) = V Then HasItrEq = True: Exit Function
Next
End Function

Function HasItrTruePrp(Itr, Prpc) As Boolean
Dim I
For Each I In Itr
    If PvzC(CvObj(I), Prpc) Then HasItrTruePrp = True: Exit Function
Next
End Function

Function IsEqNmItr(Itr, B)
IsEqNmItr = IsAySam(Itn(Itr), Itn(B))
End Function

Function AvzItrMap(Itr, Map$) As Variant()
AvzItrMap = IntozItrMap(EmpAv, Itr, Map)
End Function

Function IntozMapAy(OInto, Ay, Map$)
IntozMapAy = IntozItrMap(OInto, Itr(Ay), Map)
End Function

Function IntozItrMap(OInto, Itr, Map$)
Dim O: O = OInto
Erase O
Dim X
For Each X In Itr
    Push O, Run(Map, X)
Next
IntozItrMap = O
End Function

Sub Z_PrpVy()
Vc PrpVy(CPj.VBComponents, "CodeModule.CountOfLines")
End Sub

Function PrpVy(Itr, Prpc) As Variant()
Dim O As Object
For Each O In Itr
    Push PrpVy, PvzC(O, Prpc)
Next
End Function
Function MaxzItrPrp(Itr, Prpc)
Dim O, Obj: For Each Obj In Itr
    O = Max(O, PvzC(Obj, Prpc))
Next
MaxzItrPrp = O
End Function

Function NyzItr(Itr) As String()
NyzItr = Itn(Itr)
End Function

Function NyzItrEq(Itr, Prpc, V) As String()
Dim Obj: For Each Obj In Itr
    If PvzC(Obj, Prpc) = V Then PushI NyzItrEq, ObjNm(Obj)
Next
End Function
Function NyzOy(Oy) As String()
NyzOy = Itn(Itr(Oy))
End Function

Function VyzItrP(Itr, Prpc) As Variant()
Dim Obj: For Each Obj In Itr
    Push VyzItrP, PvzC(Obj, Prpc)
Next
End Function

Function Itn(Itr) As String()
Dim I
For Each I In Itr
    PushI Itn, ObjNm(I)
Next
End Function

Function AllTrue(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If Not P.Pred(I) Then Exit Function
Next
AllTrue = True
End Function

Function HasFalse(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If Not P.Pred(I) Then HasFalse = True: Exit Function
Next
End Function

Function HasTruePrp(Itr, Prpc) As Boolean
Dim I: For Each I In Itr
    If PvzC(I, Prpc) Then HasTruePrp = True: Exit Function
Next
End Function

Function HasTrue(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If P.Pred(I) Then HasTrue = True: Exit Function
Next
End Function

Function SyzItrPrp(Itr, P) As String()
SyzItrPrp = IntozItrPrp(EmpSy, Itr, P)
End Function

Function AvzItrPrp(Itr, P$) As Variant()
AvzItrPrp = IntozItrPrp(EmpAv, Itr, P)
End Function

Function IntozIwPredTrue(Into, Itr, P As IPred)
IntozIwPredTrue = ResiU(Into)
Dim Obj: For Each Obj In Itr
    If P.Pred(Obj) Then
        Push IntozIwPredTrue, Obj
    End If
Next
End Function

Function IntozIwEq(Into, Itr, Prpc, V)
IntozIwEq = ResiU(Into)
Dim Obj: For Each Obj In Itr
    If PvzC(Obj, Prpc) = V Then PushObj IntozIwEq, Obj
Next
End Function

Function Into(OInto, Itr)
Into = ResiU(OInto)
Dim I
For Each I In Itr
    Push Into, I
Next
End Function
Function IntozItrPrp(Into, Itr, Prpc)
IntozItrPrp = ResiU(Into)
Dim Obj: For Each Obj In Itr
    Push IntozItrPrp, PvzC(Obj, Prpc)
Next
End Function

Function IwEq(Itr, Prpc, V)
IwEq = ItrClnAy(Itr)
Dim Obj: For Each Obj In Itr
    If PvzC(Obj, Prpc) = V Then PushObj IwEq, Obj
Next
IwEq = Obj
End Function

Function IwPrpTrue(Itr, TruePrpc)
IwPrpTrue = ItrClnAy(Itr)
Dim Obj: For Each Obj In Itr
    If PvzC(Obj, TruePrpc) Then
        Push IwPrpTrue, Obj
    End If
Next
End Function

Function IwNm(Itr, B As WhNm)
IwNm = ItrClnAy(Itr)
Dim O
For Each O In Itr
    If HitNm(ObjNm(O), B) Then
        Push IwNm, O
    End If
Next
End Function


Function NIwEq&(Itr, Prpc, V)
Dim O&, Obj: For Each Obj In Itr
    If PvzC(Obj, Prpc) = V Then O = O + 1
Next
NIwEq = O
End Function

Function PrpNy(Itr) As String()
PrpNy = Itn(Itr.Properties)
End Function

Function IntozItrP(OInto, Itr, Prpc, Optional ThwEr As EmThw) As String()
Dim O: O = OInto
Dim Obj As Object
For Each Obj In Itr
    Push O, PvzC(Obj, Prpc)
Next
IntozItrP = O
End Function
Function SyzItrP(Itr, Prpc) As String()
SyzItrP = IntozItrP(EmpSy, Itr, Prpc)
End Function
