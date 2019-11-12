Attribute VB_Name = "MxSyAdd"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxSyAdd."

Function AddOsy(Osy, MsgStrOrSy)
If Not IsSy(Osy) Then Exit Function
AddOsy = AddSy(CvSy(Osy), CvSy(MsgStrOrSy))
End Function

Function AddSy(A$(), b$()) As String()
AddSy = A
PushIAy AddSy, b
End Function

Function AddSyItm(A$(), Itm$) As String()
AddSyItm = A
PushI AddSyItm, Itm
End Function

Sub Z_GpAy()
Dim Ay(), N%
GoSub T0
Exit Sub
T0:
    Ay = Array(1, 2, 3, 4, 5, 6)
    Ept = Array(Array(1, 2, 3, 4, 5), Array(6))
    N = 5
    GoTo Tst
Tst:
    Act = GpAy(Ay, N%)
    C
    Return
End Sub

Function GpAy(Ay, N%) As Variant()
Dim NEle&: NEle = Si(Ay): If NEle = 0 Then Exit Function
Dim Emp: Emp = Ay: Erase Emp
Dim M: M = Emp
Dim V, GpI%, Ix%: For Each V In Itr(Ay)
    PushI M, V
    GpI = GpI + 1
    If GpI = N Then
        GpI = 0
        PushI GpAy, M
        M = Emp
    End If
Next
If Si(M) > 0 Then PushI GpAy, M
End Function

Function AddItm(Ay, Itm)
Dim O: O = Ay
PushI O, Itm
AddItm = O
End Function

Function AddAy(AyA, AyB)
AddAy = AyA
PushAy AddAy, AyB
End Function

Function AddAv(A(), b()) As Variant()
AddAv = A
PushAy AddAv, b
End Function

Function VbTyNyzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI VbTyNyzAy, TypeName(I)
Next
End Function

Function AddAyAp(Ay, ParamArray Itm_or_AyAp())
Const CSub$ = CMod & "AddAyAp"
Dim Av(): Av = Itm_or_AyAp
If Not IsArray(Ay) Then Thw CSub, "Ay must be array", "Ay-TypeName", TypeName(Ay)
AddAyAp = Ay
Dim I: For Each I In Av
    If IsArray(I) Then
        PushIAy AddAyAp, I
    Else
        PushI AddAyAp, I
    End If
Next
End Function

Function MapAy(Ay, MapFun$) As Variant()
Dim X
For Each X In Itr(Ay)
    Push MapAy, Run(MapFun, X)
Next
End Function

Function AddEle(Ay, Ele)
AddEle = Ay
Push AddEle, Ele
End Function

Function AddEleS(Sy$(), Itm$) As String()
AddEleS = AddEle(Sy, Itm)
End Function

Function IncAy(Ay, Optional N& = 1)
IncAy = ResiU(Ay)
Dim X
For Each X In Itr(Ay)
    PushI IncAy, X + N
Next
End Function

Sub Z_AddAy()
Dim Ay1(), Ay2()
GoSub T1
Exit Sub
T1:
    Ay1 = Array(1, 2, 2, 2, 4, 5)
    Ay2 = Array(2, 2)
    Ept = Array(1, 2, 2, 2, 4, 5, 2, 2)
    GoTo Tst
Tst:
    Act = AddAy(Ay1, Ay2)
    C
    Return
End Sub

Sub Z_AmAddPfx()
Dim Sy$(), Pfx$
GoSub T1
Exit Sub
T1:
    Sy = SyzSS("1 2 3 4")
    Pfx = "* "
    Ept = SyzAp("* 1", "* 2", "* 3", "* 4")
    GoTo Tst
Tst:
    Act = AmAddPfx(Sy, Pfx)
    C
    Return
End Sub

Sub Z_AmAddPfxS()
Dim Sy$(), Act$(), Sfx$, Pfx$, Exp$()
Sy = SyzAp(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = SyzAp("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AmAddPfxS(Sy, Pfx, Sfx)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function TabSy(Sy$()) As String()
TabSy = AmAddPfx(Sy, vbTab)
End Function

Sub Z_AmAddSfx()
Dim Sy$(), Sfx$
Sy = SyzSS("1 2 3 4")
Sfx = "#"
Ept = SyzSS("1# 2# 3# 4#")
GoSub Tst
Exit Sub
Tst:
    Act = AmAddSfx(Sy, Sfx)
    C
    Return
End Sub


