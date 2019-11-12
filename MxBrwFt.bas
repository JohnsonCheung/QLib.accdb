Attribute VB_Name = "MxBrwFt"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxBrwFt."

Sub VcFt(Ft)
ShellHid FmtQQ("Code.CMd ""?""", Ft)
End Sub

Sub NoteFt(Ft)
ShellMax FmtQQ("notepad.exe ""?""", Ft)
End Sub

Sub BrwFt(Ft, Optional UseVc As Boolean)
If UseVc Then
    VcFt Ft
Else
    NoteFt Ft
End If
End Sub
