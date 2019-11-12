Attribute VB_Name = "MxTyCd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxTyCd."
Const Tp_Push$ = "|" & _
"|Sub Push{n}(O As {n}s, M As {n}" & _
"|ReDim Preserve O.Ay(O.N)" & _
"|O.Ay(O.N) = M" & _
"|O.N = O.N + 1" & _
"|End Sub"

Const Tp_Pushs$ = "|" & _
"|Sub Push{n}s(O As {n}s, M As {n}s)" & _
"|Dim J&" & _
"|For J=0 To {n}.N - 1" & _
"|    Push{n} O, A.Ay(J)" & _
"|Next" & _
"|End Sub"

Const Tp_Tys$ = "|" & _
"|Type {n}s: N As Long: Ay() As {n}: End Type"

Const Tp_Sng$ = "|" & _
"|Sub Sng{n}(A As {n}) As {n}s" & _
"|Push{n} Sng{n}, A" & _
"|End Sub"

Const Tp_Add$ = "|" & _
"|Sub Add{n}(A As {n}, B As {n}) As {n}s" & _
"|Push{n} Add{n}, A" & _
"|Push{n} Add{n}, B" & _
"|End Sub"

Const Tp$ = Tp_Tys & Tp_Push & Tp_Pushs & Tp_Add & Tp_Sng

':Mtyn:      :Nm    #Module-Type-Name#
':MthPfx-Cd: :Lines                    ! always Lines of code
Function CdMty$(Mtyn)
CdMty = FmtMacro(Tp, Mtyn)
End Function

