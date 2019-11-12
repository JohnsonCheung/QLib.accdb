Attribute VB_Name = "MxPiv"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxPiv."
Type KKCntMulItmColDy
    NKK As Integer
    Dy() As Variant
End Type
    
Enum EmAgr
    EiSum
    EiCnt
    EiAvg
End Enum
Function DyBlk(A, Kix%, Gix%) As Variant()
If Si(A) = 0 Then Exit Function
Dim J%, O, K, Blk(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In A
    K = Dr(Kix)
    Gp = Dr(Gix)
    O_Ix = IxzAy(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DyBlk = O
End Function
Function KKDrIx&(KKDr, FstColIsKKDrDy)
Dim Ix&, CurKKDr
For Each CurKKDr In Itr(FstColIsKKDrDy)
    If IsEqAy(KKDr, CurKKDr) Then
        KKDrIx = Ix
        Exit Function
    End If
    Ix = Ix + 1
Next
Ix = -1
End Function
Function KKDrToItmAyDualColDy(Dy(), KKColIx&(), ItmColIx&) As Variant()
Dim Dr, Ix&, KKDr(), Itm
Dim O() 'KKDr_To_ItmAy_DualColDy
For Each Dr In Itr(Dy)
    KKDr = AwIxy(Dy, KKColIx)
    Itm = Dr(ItmColIx)
    Ix = KKDrIx(KKDr, O)
    If Ix = -1 Then
        PushI O, Array(KKDr, Array(Itm))
    Else
        O(Ix)(1) = AddAy(O(Ix)(1), Itm)
    End If
Next
KKDrToItmAyDualColDy = O
End Function
Function KKCntMulItmColDy(Dy(), KKColIx&(), ItmColIx&) As Variant()
Dim A(): A = KKDrToItmAyDualColDy(Dy, KKColIx, ItmColIx)
KKCntMulItmColDy = KKCntMulItmColDyD(A)
End Function
Function KKCntMulItmColDyD(KKDrToItmAyDualColDy()) As Variant()

End Function
Function GpDic(A As Drs, KK$, G$) As Dictionary
Dim Fny$()
Dim KeyIxy&(), Gix%
    Fny = TermAy(KK)
    KeyIxy = Ixy(A.Fny, Fny)
    PushI Fny, G & "_Gp"
    Gix = IxzAy(Fny, G)
Set GpDic = GRxyzCyDic(A.Dy, KeyIxy, Gix)
End Function
Function DyoDotLyz2Col(DotLy$()) As Variant()
Dim O(), I, S$
For Each I In Itr(DotLy)
    S = I
    With Brk1(S, ".")
       Push O, Sy(.S1, .S2)
   End With
Next
DyoDotLyz2Col = O
End Function

Function DyoDotLy(DotLy$()) As Variant()
Dim I
For Each I In Itr(DotLy)
    PushI DyoDotLy, SplitDot(I)
Next
End Function
Function DyoDotLyzTwoCol(DotLy$()) As Variant()
Dim I: For Each I In Itr(DotLy)
    With Brk1Dot(I)
        PushI DyoDotLyzTwoCol, Array(.S1, .S2)
    End With
Next
End Function

Function DyoLyWithColon(LyWithColon$()) As Variant()
Dim I: For Each I In Itr(LyWithColon)
    PushI DyoLyWithColon, SplitColon(I)
Next
End Function

Function GRxyzCyDic(A, KeyIxy, G) As Dictionary
Const CSub$ = CMod & "GRxyzCyDic"
'If K < 0 Or G < 0 Then
'    Thw CSub, "K-Idx and G-Idx should both >= 0", "K-Idx G-Idx", K, G
'End If
'Dim Dr, U&, O As New Dictionary, KK, GG, Ay()
'U = UB(A): If U = -1 Then Exit Function
'For Each Dr In A
'    KK = Dr(K)
'    GG = Dr(G)
'    If O.Exists(KK) Then
'        Ay = O(KK)
'        PushI Ay, GG
'        O(KK) = Ay
'    Else
'        O.Add KK, Array(GG)
'    End If
'Next
'Set GRxyzCyDic = O
End Function

Function DrszFbt(Fb, T) As Drs
DrszFbt = DrszT(Db(Fb), T)
End Function

Function KE24Drs() As Drs
KE24Drs = DrszFbt(SampFbzDutyDta, "KE24")
End Function
