Attribute VB_Name = "MxVbLg"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbLg."

Sub LgrBrw()
BrwFt LgrFt
End Sub

Property Get LgrFilNo%()
LgrFilNo = FnoA(LgrFt)
End Property

Property Get LgrFt$()
LgrFt = LgrPth & "Log.txt"
End Property

Sub LgrLg(Msg$)
Dim F%: F = LgrFilNo
Print #F, NowStr & " " & Msg
If LgrFilNo = 0 Then Close #F
End Sub

Property Get LgrPth$()
Dim O$:
'O = WrkPth: EnsPth O
O = O & "Log\": EnsPth O
LgrPth = O
End Property
