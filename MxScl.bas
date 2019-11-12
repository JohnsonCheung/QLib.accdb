Attribute VB_Name = "MxScl"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxScl."

Sub AsgSclNN(Scl$, NN$, ParamArray OAp())
Const CSub$ = CMod & "AsgSclNN"
Dim Av(): Av = OAp
Dim V, Ny$(), I, J%
Ny = TermAy(NN)
If Si(Ny) <> Si(Av) Then Stop
For Each I In Itr(AeEmpEle(AmTrim(SplitSemi(Scl))))
'    V = SclItm_V(I, Ny)
    Select Case True
    Case IsByt(V) And (V = 1 Or V = 2)
    Case IsBool(V) Or IsStr(V): OAp(J) = V
    Case Else: Thw CSub, "Program error in SclItm_V.  It should return one of (Byt1,Byt2,Bool,Str)", "[But now it returns]", TypeName(V)
    End Select
    J = J + 1
Next
End Sub

Function ChkSclNN(A$, Ny0) As String()
Const CSub$ = CMod & "ChkSclNN"
Dim V, Ny$(), I, Er1$(), Er2$()
Ny = TermAy(Ny0)
For Each I In Itr(AeEmpEle(AmTrim(SplitSemi(A))))
'    V = SclItm_V(I, Ny)
    Select Case True
    Case IsByt(V) And V = 1: Push Er1, I
    Case IsByt(V) And V = 2: Push Er2, I
    Case IsBool(V) Or IsStr(V)
    Case Else: Thw CSub, "Program error in SclItm.  It should return (Byt1,Byt2,Bool,Str), but now it returns [Ty]", TypeName(V)
    End Select
Next
Dim O$()
    If Si(Er1) > 0 Then
        O = LyzMsgNap("There are [invalid-SclNy] in given [scl] under these [valid-SclNy].", "Er Ny", JnSpc(Er1), A, JnSpc(Ny))
    End If
    If Si(Er2) > 0 Then
        PushAy O, LyzMsgNap("[Itm] of [Scl] has [valid-SclNy], but it is not one of SclNy nor it has '='", "Er Scl Valid-SclNy", Er2, A, Ny)
    End If
ChkSclNN = O
End Function

Function SclItm_V(A$, Ny$())
'Return Byt1 if Pfx of A not in Ny
'Return True If A = One Of Ny
'Return Byt2 if Pfx of A is in Ny, but not Eq one Ny and Don't have =
If HasEle(Ny, A) Then SclItm_V = True: Exit Function
'If Not HasStrPfxSy(A, Ny) Then SclItm_V = CByte(1): Exit Function
If Not HasSubStr(A, "=") Then SclItm_V = CByte(2): Exit Function
SclItm_V = Trim(Aft(A, "="))
End Function

Function ShfScl$(OStr$)
AsgBrk1 OStr, ";", ShfScl, OStr
End Function



Function ShfVy(OLin, Lblss$) As Variant() ' _
'Return Ay, which is _
'   Same sz as Lblss-cnt _
'   Ay-ele will be either string of boolean _
'   Each element is corresponding to terms-lblss _
'Update OLin _
'   if the term match, it will removed from OLin _
'Lblss: is in *LLL ?LLL or LLL _
'   *LLL is always first at beginning, mean the value OLin has not lbl _
'   ?LLL means the value is in OLin is using LLL => it is true, _
'   LLL  means the value in OLin is LLL=VVV _
'OLin is _
'   VVV VVV=LLL [VVV=L L]
Dim L$, I, Ay$(), O()
Ay = TermAy(OLin)
For Each I In Itr(SyzSS(Lblss))
    L = I
    Select Case FstChr(L)
    Case "*":
        Select Case SndChr(L)
        Case "?":  If Si(Ay) = 0 Then Thw CSub, "Must BoolLbl of Lblss not found in OLin", "Must-Bool-Lbl OLin Lblss", L, OLin, Lblss
                   PushI O, CBool(ShfFstEle(Ay))
        Case Else: PushI O, ShfFstEle(Ay)
        End Select
    Case "?": PushI O, ShfBool(Ay, L)
    Case Else: PushI O, ShfTxt(Ay, L)
    End Select
Next
ShfVy = O
End Function

Function ShfTxtOpt(OAy$(), Lbl) As StrOpt
If Si(OAy) = 0 Then Exit Function
Dim S$: S = ShfTxt(OAy, Lbl)
If S = "" Then ShfTxtOpt = SomStr(S)
End Function

Function ShfBool(OAy$(), Lbl$)
If Si(OAy) = 0 Then Exit Function
Dim J&, L$, Ay$()
Ay = OAy
L = RmvFstChr(Lbl)
For J = 0 To UB(Ay)
    If Ay(J) = L Then
        ShfBool = True
        OAy = AeEleAt(Ay, J)
        Exit Function
    End If
Next
End Function

Function ShfTxt(OAy$(), Lbl)
If Si(OAy) = 0 Then Exit Function
'Return either string or ""
Dim I, S$, J&, Ay$()
Ay = OAy
For Each I In Itr(Ay)
    S = I
    With Brk2(S, "=")
        If .S1 = Lbl Then
            ShfTxt = .S2
            OAy = AeEleAt(Ay, J)
            Exit Function
        End If
    End With
    J = J + 1
Next
End Function

Sub Z_ShfVy()
GoSub T0
'GoSub T1
'GoSub T2
'GoSub T3
Exit Sub
Dim Lin, Lblss$, Act(), Ept()
T0:
    Lin = "Loc Txt Req Dft=ABC AlwZLen [VTxt=Loc cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
    Lblss = EleLblss ' "*Fld *Ty ?Req ?AlwZLen Dft VTxt VRul TxtSz Expr"
    Ept = Array("Loc", "Txt", True, True, "ABC", "Loc cannot be blank", "IsNull([Loc]) or Trim(Loc)=''", "", "")
    GoTo Tst
T1:
    Lin = "1 Req"
    Lblss = "*XX ?Req"
    Ept = Array("1", True)
    GoTo Tst
T2:
    Lin = "A B C=123 D=XYZ"
    Lblss = "?B"
    Ept = Array(True)
    GoTo Tst
T3:
    Lin = "Txt VTxt=XYZ [Dft=A 1] VRul=123 Req"
    Lblss = "*Fld *Ty ?Req ?AlwZLen Dft VTxt VRul"
    Ept = Array("Txt", True, False, "A 1", "XYZ", "123")
    GoTo Tst
Tst:
    Act = ShfVy(Lin, Lblss)
    C
    Return
End Sub
