Attribute VB_Name = "MxDro3Cnst"
Option Explicit
Option Compare Text
':CNsv: :S #Cnst-CNs-Value# ! the string bet-DblQ of CnstLin-CNs of a Md
':CModv: :S #Cnst-CMod-Value# ! the string aft-rmv-sfx-[.] of bet-DblQ of CnstLin-CMod of a Md
':CLibv: :S #Cnst-CLib-Value# ! the string aft-rmv-sfx-[.] of bet-DblQ of CnstLin-CLib of a Md
Const CMod$ = CLib & "MxDro3Cnst."
Public Const FFoCMod$ = "CLibv CNsv CModv"

Function FoCMod() As String()
FoCMod = SyzSS(FFoCMod)
End Function

Function DroCMod(M As CodeModule)
Dim D$(): D = DclzM(M)
DroCMod = Array(CLibv(D), CNsv(D), CModv(D))
End Function

