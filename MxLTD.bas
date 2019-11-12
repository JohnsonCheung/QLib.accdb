Attribute VB_Name = "MxLTD"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxLTD."
Public Const FFoLTD$ = "L T1 Dta"
Public Const FFoLTTD$ = "L T1 T2 Dta"
Type TDoLTD: D As Drs: End Type

Function DoErLTD(DoLTD As Drs, T1ss$) As Drs
Dim T1Ay$(): T1Ay = SyzSS(T1ss)
Dim ODy()
    Dim Dr: For Each Dr In Itr(DoLTD.Dy)
        If Not HasEle(T1Ay, Dr(0)) Then
            PushI ODy, Dr
        End If
    Next
DoErLTD = Drs(DoLTD.Fny, ODy)
End Function

Function TDoLTDzInd(IndentSrc$()) As TDoLTD
TDoLTDzInd = TDoLTDzH(TDoLTDH(IndentSrc))
End Function

Function DoLTD(Src$()) As Drs
DoLTD = DrszFF(FFoLTD, DyoLTD(Src))
End Function

Function DoLTTD(Src$()) As Drs
DoLTTD = DrszFF(FFoLTTD, DyoLTTD(Src))
End Function

Function TDoLTDzH(A As TDoLTDH) As TDoLTD
TDoLTDzH.D = DwFalseExl(A.D, "IsHdr")
End Function

Private Function DyoLTTD(Src$()) As Variant()
'Ret :Dy-L-T1-T2-Dta
Dim L&, Dta$, T1$, T2$
Dim Lin: For Each Lin In Itr(Src)
    L = L + 1
    If Fst2Chr(LTrim(L)) <> "--" Then
        AsgTTRst Lin, T1, T2, Dta
        PushI DyoLTTD, Array(L, T1, T2, Dta)
    End If
Next
End Function

Private Function DyoLTD(Src$()) As Variant()
'Ret:: Dy{L T1 Dta}
Dim L&, Dta$, T1$, Lin
For Each Lin In Itr(Src)
    L = L + 1
    If Fst2Chr(LTrim(L)) = "--" Then GoTo X
    T1 = T1zS(Lin)
    Dta = RmvT1(Lin)
    PushI DyoLTD, Array(L, T1, Dta)
X:
Next
End Function

Function EoLTDTy(D As TDoLTD, Tyss$) As String()
Dim TyAy$(): TyAy = SyzSS(Tyss)
Dim T$, L&
Dim Dr: For Each Dr In Itr(D.D.Dy)
    T = Dr(1)
    If Not HasEle(TyAy, T) Then
        L = Dr(0)
        PushI EoLTDTy, MoTyEr(L, T, Tyss)
    End If
Next
End Function

Function MoTyEr$(L&, T$, Tyss$)
':MoTyEr: :MsgStr #Msg-Of-TyEr#
MoTyEr = FmtQQ("L#(?) has Ty(?) which is not in valid Tyss(?)", L, T, Tyss)
End Function
