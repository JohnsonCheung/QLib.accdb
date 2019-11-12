Attribute VB_Name = "MxCmlnss"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCmlnss."
Public Const FFoCml$ = "Nm C1"
Function Cmlnss$(S)
':Cmlnss: :S #Cml<name-ss>#
Cmlnss = S & " " & Cmlss(S)
End Function

Function DoCmlnss(Ny$()) As Drs
Dim Dy(): Dy = DyoCmlnss(Ny)
Dim N%: N = NColzDy(Dy)
Dim R$: R = Rest(N)
DoCmlnss = DrszFF(FFoCml & R, DyoCmlnss(Ny))
End Function
Private Function Rest$(N%)
'Ret " C2...C<N>"
Dim O$()
Dim J%: For J = 2 To N - 1
    PushI O, " C" & J
Next
Rest = JnSpc(O)
End Function
Function DyoCmlnss(Ny$()) As Variant()
DyoCmlnss = DyzSSAy(CmlnssAy(Ny))
End Function
Function CmlnssAy(Ny$()) As String()
Dim I: For Each I In Itr(Ny)
    PushI CmlnssAy, Cmlnss(I)
Next
End Function
