Attribute VB_Name = "MxItrm"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxItrm."
Function ImAddSfx(Itr, Sfx$) As String()
Dim V: For Each V In Itr
    Push ImAddSfx, V & Sfx
Next
End Function

Function ImAddPfx(Itr, Pfx$) As String()
Dim V: For Each V In Itr
    Push ImAddPfx, Pfx & V
Next
End Function

