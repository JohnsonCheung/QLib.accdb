Attribute VB_Name = "MxJmp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxJmp."

Sub JmpNxt()

End Sub

Sub JmpCmpn(Cmpn$)
Dim C As VBIDE.CodePane: Set C = PnezCmpn(Cmpn)
If IsNothing(C) Then Debug.Print "No such WinOfCmpNm": Exit Sub
C.Show
End Sub

