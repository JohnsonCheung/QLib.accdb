Attribute VB_Name = "MxIxAy"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxIxAy."
Type NumPos
    Num As Long
    Pos As Long
End Type

Sub Z_AsgIx()
Dim Drs As Drs, FF$, A%, b%, C%, EA%, EB%, Ec%
GoSub T1
Exit Sub
T1:
    Drs.Fny = SyzSS("A B C")
    FF = "C B A"
    EA = 0
    EB = 1
    Ec = 2
    GoTo Tst
Tst:
    AsgIx Drs, FF, A, b, C
    Debug.Print A = EA
    Debug.Print b = EB
    Debug.Print C = Ec
    Return
End Sub

Sub AsgIx(A As Drs, FF$, ParamArray OIxAp())
Dim F, J%, I&
For Each F In SyzSS(FF)
    I = IxzAy(A.Fny, F, ThwEr:=EiThwEr)
    OIxAp(J) = IxzAy(A.Fny, F)
    J = J + 1
Next
End Sub

Function Ele(Ay, Ix)
If IsBet(Si(Ay), 0, Ix) Then Ele = Ay(Ix)
End Function
Function IxzAy&(Ay, Itm, Optional FmIx& = 0, Optional ThwEr As EmThw)
Dim J&
For J = FmIx To UB(Ay)
    If Ay(J) = Itm Then IxzAy = J: Exit Function
Next
If ThwEr = EiThwEr Then
    Thw CSub, "Itm not found in Ay", "Itm Si(Ay) Ay", Itm, Si(Ay), Ay
End If
IxzAy = -1
End Function

Function IxyzU(U&) As Long()
Dim O&(), J&
ReDim O(U)
For J = 0 To U
    O(J) = J
Next
IxyzU = O
End Function

Function GrpEndIxyOfEmp(Ay) As Long()
Dim J&, Fst As Boolean, U&
U = UB(Ay)
For J = 1 To U
    If Not IsEmpty(Ay(J)) Then
        PushI GrpEndIxyOfEmp, J - 1
    End If
Next
PushI GrpEndIxyOfEmp, U
End Function

Function IxyzEle(Ay, Ele) As Long()
Dim J&, V
For Each V In Itr(Ay)
    If V = Ele Then PushI IxyzEle, J
    J = J + 1
Next
End Function
Function IxyzAlwE(Ay, SubAy) As Long()
Dim I, HasNegIx As Boolean, Ix&
For Each I In Itr(SubAy)
    Ix = IxzAy(Ay, I, , EiNoThw)
    PushI IxyzAlwE, Ix
Next
End Function
Function CnyzSubFny(Fny$(), SubFny$()) As Integer()
Dim F: For Each F In Itr(SubFny)
    Dim I%: I = IxzAy(Fny, F)
    If I = -1 Then
        Thw CSub, "Ele in SubFny is not found in Fny", "Ele SubFny Fny", F, SubFny, Fny
    End If
    PushI CnyzSubFny, I + 1
Next
End Function
Function Ixy(Ay, SubAy) As Long()
Dim O&(): O = IxyzAlwE(Ay, SubAy)
Dim Ix: For Each Ix In Itr(O)
    If Ix = -1 Then
        Thw CSub, "Negative index", "Ay SubAy Ixy", Ay, SubAy, Ixy
    End If
Next
Ixy = O
End Function

Sub AsgItmAyIxay(Ay, Ixy&(), ParamArray OItmAp())
Dim J&
For J = 0 To UB(Ixy)
    Asg Ay(Ixy(J)), OItmAp(J)
Next
End Sub

Function IntIxy(Ay, SubAy) As Integer()
Dim J&, U&
For J = 0 To UB(SubAy)
    PushI IntIxy, IxzAy(Ay, SubAy(J))
Next
End Function
Function IxyzDup(Ay) As Long()
Dim A As Aset: Set A = AsetzAy(AwDup(Ay))
If A.IsEmp Then Exit Function
Dim J&
For J = 0 To UB(Ay)
    If A.Has(Ay(J)) Then PushI IxyzDup, J
Next
End Function
