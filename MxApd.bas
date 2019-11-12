Attribute VB_Name = "MxApd"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxApd."
Const CNs$ = "Str"
':FF: :Lin #Fldn-Spc-Sep# ! a list of Fldn has no space and separated by space.
Function PpdSpcIf$(S)
PpdSpcIf = PpdIf(S, " ")
End Function

Function AddNBzAy(Ay, Sfx$) As String()
Dim I
For Each I In Itr(Ay)
    PushI AddNBzAy, AddNB(I, Sfx)
Next
End Function

Function PpdIf$(S, Pfx$)
If S = "" Then Exit Function
PpdIf = Pfx & S
End Function


Function AddNB$(ParamArray StrAp())
'Ret : :S ! ret a str by adding each ele of @StrAp one by one, if all them is <>'' else ret blank @@
Dim Av(): Av = StrAp
Dim O$()
Dim S: For Each S In Itr(Av)
    If S = "" Then Exit Function
    PushI O, S
Next
AddNB = Jn(O)
End Function
