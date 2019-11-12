Attribute VB_Name = "MxDym"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDym."

Function DymJnDot(Dy()) As String()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DymJnDot, JnDot(Dr)
Next
End Function
