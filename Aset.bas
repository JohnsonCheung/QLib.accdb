VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Aset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "Aset."
Private Aset As New Dictionary

Function AddAset(A As Aset) As Aset
Dim O As New Aset
O.PushAset Aset
O.PushAset A
Set AddAset = O
End Function

Function Av() As Variant()
Av = AvzItr(Itms)
End Function

Sub Brw(Optional Fnn$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
B Aset.Keys, DftStr(Fnn, "Aset"), OupTy
End Sub

Function Clone() As Aset
Dim O As New Aset
O.PushItr Aset.Keys
Set Clone = O
End Function

Property Get Cnt&()
Cnt = Aset.Count
End Property

Sub Dmp()
D Aset.Keys
End Sub

Function FstItm()
Const CSub$ = CMod & "FstItm"
If IsEmp Then Thw CSub, "Given Aset is empty"
Dim I
For Each I In Itms
    Asg I, FstItm
Next
End Function

Function Has(Itm) As Boolean
Has = Aset.Exists(Itm)
End Function

Function IsEmp() As Boolean
IsEmp = Aset.Count = 0
End Function

Function IsEq(B As Aset) As Boolean
If Cnt <> B.Cnt Then Exit Function
Dim K
For Each K In Itms
    If Not B.Has(K) Then Exit Function
Next
IsEq = True
End Function

Function IsInOrdEq(B As Aset) As Boolean
IsInOrdEq = IsEqAy(Me.Av, B.Av)
End Function

Function Itms()
Itms = Aset.Keys
End Function

Function Lin()
Lin = JnSpc(Aset.Keys)
End Function

Function Minus(B As Aset) As Aset
Dim O As Aset, I
Set O = EmpAset
For Each I In Itms
    If Not B.Has(I) Then O.PushItm I
Next
Set Minus = O
End Function

Sub PushAset(A As Aset)
PushItr A.Itms
End Sub

Sub PushAy(A)
Dim I
For Each I In Itr(A)
    PushItm I
Next
End Sub

Sub PushItm(Itm)
If Not Has(Itm) Then Aset.Add Itm, Empty
End Sub

Sub PushItr(Itr, Optional NoBlnkStr As Boolean)
Dim I
If NoBlnkStr Then
    For Each I In Itr
        If I <> "" Then
            PushItm I
        End If
    Next
Else
    For Each I In Itr
        PushItm I
    Next
End If
End Sub

Function RmvItm(Itm) As Aset
If Has(Itm) Then Aset.Remove Itm
Set RmvItm = Me
End Function

Function Srt() As Aset
Set Srt = AsetzAy(AySrtQ(Itms))
End Function

Function Sy() As String()
Sy = SyzAy(Aset.Keys)
End Function

Property Get Termss()
Dim I, O$()
For Each I In Itms
    PushI O, QteSqIf(I)
Next
Termss = JnSpc(O)
End Property
Sub Vc()
Brw OupTy:=EiOtVc
End Sub
