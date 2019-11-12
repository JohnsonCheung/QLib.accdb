Attribute VB_Name = "MxAp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAp."

Function AvzAy(Ay) As Variant()
If IsAv(Ay) Then AvzAy = Ay: Exit Function
Dim I
For Each I In Itr(Ay)
    Push AvzAy, I
Next
End Function

Function Av(ParamArray ApOf_Itm_Or_Ay()) As Variant()
Dim Av1(): Av = ApOf_Itm_Or_Ay
Av1 = ApOf_Itm_Or_Ay
Av = AvzAyOfItmOrAy(Av1)
End Function

Function AvzAyOfItmOrAy(AyOfItmOrAy) As Variant()
Dim V: For Each V In Itr(AyOfItmOrAy)
    If IsArray(V) Then
        PushIAy AvzAyOfItmOrAy, V
    Else
        PushI AvzAyOfItmOrAy, V
    End If
Next
End Function

Function AvzAp(ParamArray ApOf_Itm_Or_Ay()) As Variant()
Dim Av(): Av = ApOf_Itm_Or_Ay
Av = ApOf_Itm_Or_Ay
AvzAp = AvzAyOfItmOrAy(Av)
End Function

Function SyzAv(AvOf_Itm_or_Ay()) As String()
Dim I: For Each I In Itr(AvOf_Itm_or_Ay)
    If IsArray(I) Then
        PushIAy SyzAv, I
    Else
        PushNB SyzAv, I
    End If
Next
End Function

Function SyzAvNB(AvOf_Itm_or_Ay()) As String()
Dim I: For Each I In Itr(AvOf_Itm_or_Ay)
    If IsArray(I) Then
        PushNBAy SyzAvNB, I
    Else
        PushNB SyzAvNB, I
    End If
Next
End Function

Function SyzAp(ParamArray ApOf_Itm_Or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay
SyzAp = SyzAv(Av)
End Function

Function SyzApNB(ParamArray ApOf_Itm_Or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay
SyzApNB = SyzAvNB(Av)
End Function

Function Sy(ParamArray ApOf_Itm_Or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay
Sy = SyzAv(Av)
End Function

Function DteAy(ParamArray Ap()) As Date()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
DteAy = IntozAy(DteAy, Av)
End Function

Function IntAyzLngAy(LngAy&()) As Integer()
Dim I
For Each I In Itr(LngAy)
    PushI IntAyzLngAy, I
Next
End Function

Function IntAyzSS(IntSS$) As Integer()
Dim I
For Each I In Itr(SyzSS(IntSS))
    PushI IntAyzSS, I
Next
End Function

Function IntAy(ParamArray Ap()) As Integer()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
IntAy = IntozAy(EmpIntAy, Av)
End Function

Function LngAy(ParamArray Ap()) As Long()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
LngAy = IntozAy(EmpLngAy, Av)
End Function

Function SngAy(ParamArray Ap()) As Single()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
SngAy = IntozAy(SngAy, Av)
End Function

Function SyNB(ParamArray S_or_Sy()) As String()
Dim Av(): Av = S_or_Sy
Dim I
For Each I In Av
    Select Case True
    Case IsArray(I): PushNBAy SyNB, I
    Case Else: PushNB SyNB, I
    End Select
Next
End Function

Function SyzAy(Ay) As String()
If IsSy(Ay) Then SyzAy = Ay: Exit Function
SyzAy = IntozAy(EmpSy, Ay)
End Function
