Attribute VB_Name = "MxCprDic"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxCprDic."
Type CprDic
    Nm1 As String
    Nm2 As String
    AExcess As Dictionary
    BExcess As Dictionary
    ADif As Dictionary
    BDif As Dictionary
    Sam As Dictionary
End Type

Function FmtCprDic(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd", Optional ExlSam As Boolean) As String()
FmtCprDic = FmtCprDiczCpr(CprDic(A, B, Nm1, Nm2), ExlSam)
End Function

Function FmtCprDiczCpr(A As CprDic, Optional ExlSam As Boolean) As String()
Dim O$()
With A
    O = AddAyAp( _
        FmtExcess(.AExcess, .Nm1), _
        FmtExcess(.BExcess, .Nm2), _
        FmtDif(.ADif, .BDif))
End With
If Not ExlSam Then
    O = AddAy(O, FmtSam(A.Sam))
End If
FmtCprDiczCpr = O
End Function

Function CprDic(A As Dictionary, B As Dictionary, Nm1$, Nm2$) As CprDic
With CprDic
    .Nm1 = Nm1
    .Nm2 = Nm2
    Set .AExcess = MinusDic(A, B)
    Set .BExcess = MinusDic(B, A)
    Set .Sam = SamKV(A, B)
    AsgADifBDif A, B, .ADif, .BDif
End With
End Function

Sub BrwCprDic(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd", Optional ExlSam As Boolean)
BrwAy FmtCprDic(A, B, Nm1, Nm2)
End Sub

Function SamKV(A As Dictionary, B As Dictionary) As Dictionary
Set SamKV = New Dictionary
If A.Count = 0 Or B.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            SamKV.Add K, A(K)
        End If
    End If
Next
End Function

Sub AsgADifBDif(A As Dictionary, B As Dictionary, _
    OADif As Dictionary, OBDif As Dictionary)
Dim K
Set OADif = New Dictionary
Set OBDif = New Dictionary
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            OADif.Add K, A(K)
            OBDif.Add K, B(K)
        End If
    End If
Next
End Sub

Function FmtDif(A As Dictionary, B As Dictionary) As String()
If A.Count <> B.Count Then Thw CSub, "Dic A & B should have same size", "Dic-A-Si Dic-B-Si", A.Count, B.Count
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S As S12s, KK$
For Each K In A
    KK = K
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & ULinzLines(KK) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & ULinzLines(KK) & vbCrLf & B(K)
    PushS12 S, S12(S1, S2)
Next
FmtDif = FmtS12s(S, FF:="")
End Function

Function FmtExcess(A As Dictionary, Nm$) As String()
If A.Count = 0 Then Exit Function
Dim K, S1$, S2$, S As S12s
S2 = "!" & "Er Excess (" & Nm & ")"
For Each K In A.Keys
    S1 = ULinzLines(CStr(K))
    S2 = A(K)
    PushS12 S, S12(S1, S2)
Next
PushAy FmtExcess, FmtS12s(S, FF:="Exccess " & Nm)
End Function

Function FmtSam(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S As S12s, KK$
For Each K In A.Keys
    KK = K
    PushS12 S, S12("*Same", K & vbCrLf & ULinzLines(KK) & vbCrLf & A(K))
Next
FmtSam = FmtS12s(S)
End Function

Sub Z_BrwCprDic()
Dim A As Dictionary, B As Dictionary
Set A = DiczVbl("X AA|A BBB|A Lines1|A Line3|B Line1|B line2|B line3..")
Set B = DiczVbl("X AA|C Line|D Line1|D line2|B Line1|B line2|B line3|B Line4")
BrwCprDic A, B
End Sub

