Attribute VB_Name = "MxAySrt"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAySrt."
Enum EmOrd
    EiAsc
    EiDes
End Enum
Private A_Ay
Function SrtLines$(A$)
SrtLines = JnCrLf(AySrt(SplitCrLf(A)))
End Function

Function IsSrtdzAy(Ay) As Boolean
Dim J&
For J = 0 To UB(Ay) - 1
   If Ay(J) > Ay(J + 1) Then Exit Function
Next
IsSrtdzAy = True
End Function

Sub Z_AySrtQ()
Dim Ay, IsDes As Boolean
GoSub T0
GoSub T1
Exit Sub
T0:
    Ay = Array(1, 2, 3, 4, 0, 1, 1, 5)
    IsDes = False
    Ept = Array(0, 1, 1, 1, 2, 3, 4, 5)
    GoTo Tst
T1:
    Ay = Array(1, 2, 4, 87, 4, 2)
    IsDes = True
    Ept = Array(87, 4, 4, 2, 2, 1)
    GoTo Tst
Tst:
    Act = AySrtQ(Ay, IsDes:=IsDes)
    C
    Return
End Sub



Function AySrtQ(Ay, Optional IsDes As Boolean)
If Si(Ay) = 0 Then AySrtQ = Ay: Exit Function
A_Ay = Ay  ' Put it is A_Ay, which is untouch, only sort the %Ixy
Dim Ixy&(): Ixy = LngSeq(0, UB(Ay))
Ixy = AySrtQ__Srt(Ixy)  'Srt the %Ixy
If IsDes Then
    Ixy = AyRevI(Ixy)
End If
AySrtQ = AwIxy(Ay, Ixy)
End Function

Function AySrtQ__Srt(Ixy&()) As Long()
'Ret : a sorted ix of @Ixy
Dim O&()
    Select Case Si(Ixy)
    Case 0
    Case 1: O = Ixy
    Case 2:
        If AySrtQ__IsLE(Ixy(0), Ixy(1)) Then
            O = Ixy
        Else
            PushI O, Ixy(1)
            PushI O, Ixy(0)
        End If
    Case Else
        Dim L&():   L = Ixy
        Dim P&:   P = Pop(L)
        Dim A&(): A = AySrtQ__LE(L, P)
        Dim B&(): B = AySrtQ__GT(L, P)
                  A = AySrtQ__Srt(A) 'Srt-it
                  B = AySrtQ__Srt(B)
        PushIAy O, A
          PushI O, P
        PushIAy O, B
    End Select
AySrtQ__Srt = O
End Function
Function AySrtQ__IsLE(A, B&) As Boolean
AySrtQ__IsLE = A_Ay(A) <= A_Ay(B)
End Function
Function AySrtQ__LE(Ixy&(), P&) As Long()
'@Ixy : Ix to be selected
'@P   : the Pivot-Ix to select those ix in @Ixy
'Ret : a subset of @Ixy so that each the ret ix is LE than the Pivot-@P
Dim I: For Each I In Itr(Ixy)
    If AySrtQ__IsLE(I, P) Then PushI AySrtQ__LE, I  ' If the running-Ix-%I IsGT pivot-Ix-@P, push it to ret
Next
End Function
Function AySrtQ__GT(Ixy&(), P&) As Long()
'@Ixy : Ix to be selected
'@P   : the Pivot-Ix to select those ix in @Ixy
'Ret : a subset of @Ixy so that each the ret ix is GT than the Pivot-@P
Dim I: For Each I In Itr(Ixy)
    If Not AySrtQ__IsLE(I, P) Then PushI AySrtQ__GT, I  ' If the running-Ix-%I IsGT pivot-Ix-@P, push it to ret
Next
End Function
    
    


        
    


Sub Z_AySrtByAy()
Dim Ay, ByAy
Ay = Array(1, 2, 3, 4)
ByAy = Array(3, 4)
Ept = Array(3, 4, 1, 2)
GoSub Tst
Exit Sub
Tst:
    Act = AySrtByAy(Ay, ByAy)
    C
    Return
End Sub

Function AySrtByAy(Ay, ByAy)
Dim O: O = ResiU(Ay)
Dim I
For Each I In ByAy
    If HasEle(Ay, I) Then PushI O, I
Next
PushIAy O, AyMinus(Ay, O)
AySrtByAy = O
End Function

Function AySrt(Ay, Optional Des As Boolean)
If Si(Ay) = 0 Then AySrt = Ay: Exit Function
Dim Ix&, V, J&
Dim O: O = Ay: Erase O
Push O, Ay(0)
For J = 1 To UB(Ay)
    O = InsEle(O, Ay(J), AySrt__Ix(O, Ay(J)))
Next
If Des Then
    AySrt = AyRev(O)
Else
    AySrt = O
End If
End Function

Function AySrt__Ix&(A, V)
Dim I, O&
For Each I In A
    If V < I Then AySrt__Ix = O: Exit Function
    O = O + 1
Next
AySrt__Ix = O
End Function

Function IxyzSrtAy(Ay, Optional Des As Boolean) As Long()
If Si(Ay) = 0 Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = InsEle(O, J, IxyzSrtAy_Ix(O, Ay, Ay(J), Des))
Next
IxyzSrtAy = O
End Function

Sub Z_AySrt()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                   Act = AySrt(A):       ThwIf_AyabNE Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(A, True): ThwIf_AyabNE Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(A):       ThwIf_AyabNE Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin:Sub"
Push A, "Private:HasPfx:Function"
Push A, "Private:MdMthDRsFunBdyLy:Function"
Push A, "Private:SrcMthLx_ToLx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:HasPfx:Function"
Push Exp, "Private:MdMthDRsFunBdyLy:Function"
Push Exp, "Private:SrcMthLx_ToLx:Function"
Push Exp, "~Private:JnContinueLin:Sub"
Push Exp, "~~:Tst:Sub"
Act = AySrtQ(A)
ThwIf_AyabNE Exp, Act
End Sub

Function IxyzSrtAy_Ix&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then IxyzSrtAy_Ix& = O: Exit Function
        O = O + 1
    Next
    IxyzSrtAy_Ix& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then IxyzSrtAy_Ix& = O: Exit Function
    O = O + 1
Next
IxyzSrtAy_Ix& = O
End Function

Sub Z_IxyzSrtAy()
Dim A: A = Array("A", "B", "C", "D", "E")
ThwIf_AyabNE Array(0, 1, 2, 3, 4), IxyzSrtAy(A)
ThwIf_AyabNE Array(4, 3, 2, 1, 0), IxyzSrtAy(A, True)
End Sub

Function AySrtInEIxIxy&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then AySrtInEIxIxy& = O: Exit Function
        O = O + 1
    Next
    AySrtInEIxIxy& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then AySrtInEIxIxy& = O: Exit Function
    O = O + 1
Next
AySrtInEIxIxy& = O
End Function

Function DiczAddIxToKey(A As Dictionary) As Dictionary
Dim O As New Dictionary, K, J&
For Each K In A.Keys
    O.Add J & " " & K, A(K)
    J = J + 1
Next
Set DiczAddIxToKey = O
End Function

Function SrtDic(A As Dictionary, Optional IsDesc As Boolean) As Dictionary
If A.Count = 0 Then Set SrtDic = New Dictionary: Exit Function
Dim O As New Dictionary
Dim Srt: Srt = AySrtQ(A.Keys, IsDesc)
Dim K: For Each K In Srt
   O.Add K, A(K)
Next
Set SrtDic = O
End Function

Sub Z_AySrt4()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                    Act = AySrt(A):        ThwIf_NE Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(A, True): ThwIf_NE Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(A):       ThwIf_NE Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin:Sub"
Push A, "Private:HasPfx:Function"
Push A, "Private:MdMthDRsFunBdyLy:Function"
Push A, "Private:SrcMthLx_ToLx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:HasPfx:Function"
Push Exp, "Private:MdMthDRsFunBdyLy:Function"
Push Exp, "Private:SrcMthLx_ToLx:Function"
Push Exp, "~Private:JnContinueLin:Sub"
Push Exp, "~~:Tst:Sub"
Act = AySrt(A)
ThwIf_NE Exp, Act
End Sub

Sub Z_IxyzSrtAy5()
Dim A: A = Array("A", "B", "C", "D", "E")
ThwIf_NE Array(0, 1, 2, 3, 4), IxyzSrtAy(A)
ThwIf_NE Array(4, 3, 2, 1, 0), IxyzSrtAy(A, True)
End Sub


