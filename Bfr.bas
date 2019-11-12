VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bfr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "Bfr."
Private A$()

Sub Lin()
PushI A, ""
End Sub
Sub Var(Optional V)
If IsEmpty(V) Then PushI A, "": Exit Sub
PushIAy A, Fmt(V)
End Sub
Sub Box(S$, Optional C$ = "*")
PushIAy A, MxBox.Box(S, C)
End Sub

Sub ULin(S$, Optional ULinChr$ = "-")
PushI A, S
PushI A, Dup(FstChr(ULinChr), Len(S))
End Sub

Sub Brw()
BrwAy A
End Sub

Function Ly() As String()
Ly = A
End Function

Function Lines$()
Lines = JnCrLf(Ly)
End Function

Private Property Get IEr_Er() As String()
MsgBox "asdf"
End Property
