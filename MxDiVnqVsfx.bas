Attribute VB_Name = "MxDiVnqVsfx"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDiVnqVsfx."
':Vsfx: :Term #Var-Sfx# ! It is from Arg or DimItm or Fun.  It is a sht form in direct attach to a :Var or :Argn or :Mthn
':Vn:   :Nm   #Var-Nm#
Function DiVnqVsfx(MthLy$()) As Dictionary
Dim L: For Each L In Itr(MthLy)
    PushDiS12 DiVnqVsfx, S12oDimnqVsfx(DimItmAyzS(MthLy))
Next
End Function

Function S12oDimnqVsfx(DimItm) As S12
Dim L$: L = DimItm
S12oDimnqVsfx.S1 = ShfNm(L)
S12oDimnqVsfx.S2 = Vsfx(L)
End Function

Function Vsfx$(DimItm_AftNm$)
Dim S$: S = LTrim(DimItm_AftNm)
Select Case True
Case S = "": Exit Function
Case Fst2Chr(S) = "()"
    If Len(S) = 2 Then
        Vsfx = S
    Else
        S = LTrim(Mid(S, 3))
        If HasPfx(S, "As ") Then
            Vsfx = ":" & Trim(RmvPfx(S, "As ")) & "()"
        Else
            Thw CSub, "Invalid DimItm_AftNm", "When aft :() , it should be :As", "DimItm_AftNm", DimItm_AftNm
        End If
    End If
Case HasPfx(S, "As ")
    Vsfx = ":" & RmvPfx(RmvPfx(S, "As "), "New ")
Case Else
    Vsfx = S
End Select
End Function

Function S12oDimnqVsfxP() As S12s
S12oDimnqVsfxP = S12soDimnqVsfxzP(CPj)
End Function

Function S12soDimnqVsfxzP(P As VBProject) As S12s
S12soDimnqVsfxzP = S12soDimnqVsfx(DimItmAyzS(SrczP(P)))
End Function

Function S12soDimnqVsfx(DimItmAy$()) As S12s
Dim I: For Each I In Itr(DimItmAy)
    PushS12 S12soDimnqVsfx, S12oDimnqVsfx(I)
Next
End Function
