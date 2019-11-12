Attribute VB_Name = "MxDDRmk"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDDRmk."

Function HasDDRmk(Lin) As Boolean
HasDDRmk = HasSubStr(Lin, "--")
End Function

Function RmvDDRmk$(Lin)
RmvDDRmk = BefOrAll(Lin, "--", True)
End Function
Function RmvDDRmkzL(Ly$()) As String()
Dim L: For Each L In Itr(Ly)
    PushI RmvDDRmkzL, RmvDDRmk(L)
Next
End Function
