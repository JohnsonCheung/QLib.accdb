VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "DrsClass."
Private Type A
    Fny() As String
    Dy() As Variant
End Type
Private A As A
Property Get Fny() As String()
Fny = A.Fny
End Property
Property Get Dy() As Variant()
Dy = A.Dy
End Property

Sub Init(Fny$(), Dy())
A.Fny = Fny
A.Dy = Dy
End Sub

Sub IniByDrs(D As Drs)
A.Fny = D.Fny
A.Dy = D.Dy
End Sub
    
