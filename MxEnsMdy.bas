Attribute VB_Name = "MxEnsMdy"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEnsMdy."

Sub EnsPrvzNm(Mdn$, Mthn$)
'Ret : Ens a @Mthn in @Mdn as Private @@
If Not HasMd(CPj, Mdn, IsInf:=True) Then Exit Sub
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&: L = MthLnozMM(M, Mthn, IsInf:=True)

End Sub
Function EnsPrv$(MthLin)
Const CSub$ = CMod & "EnsPrv"
If Not IsLinMth(MthLin) Then Thw CSub, "Given MthLin is not MthLin", "Lin", MthLin
EnsPrv = "Private " & RmvMdy(MthLin)
End Function

Function EnsPub$(MthLin)
Const CSub$ = CMod & "EnsPub"
If Not IsLinMth(MthLin) Then Thw CSub, "Given MthLin is not MthLin", "MthLin", MthLin
EnsPub = RmvMdy(MthLin)
End Function

Sub EnsPjPrvZ()
EnsPrvZzP CPj
End Sub

Sub EnsPrvZzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsPrvZ C.CodeModule
Next
End Sub
Function Z_EnsPrv(A As Drs) As Drs

End Function
Sub EnsPrvZ(M As CodeModule, Optional Upd)
Const CmPfx$ = "X_"
Dim A As Drs: ' A = DPubZMth(M) ' L MthLin
Dim b As Drs: ' B = X_EnsPrv(A)   ' L MthLin PrvZ
Dim C As Drs: C = SelDrsAs(b, "L PrvZ:NewL MthLin:OldL")

RplLin M, C
End Sub

Function LnoAyOfPubZ(M As CodeModule) As Long()
Dim L, J&
For Each L In Itr(Src(M))
    J = J + 1
    'If IsMthLinzPub(L) Then
        PushI LnoAyOfPubZ, J
    'End If
Next
End Function

Function LnoItrOfPubZ(M As CodeModule)
Asg Itr(LnoAyOfPubZ(M)), LnoItrOfPubZ
End Function

Function EnsMdy$(OldMthLin, ShtMdy$)
Const CSub$ = CMod & "EnsMdy"
Dim L$: L = RmvMdy(OldMthLin)
    Select Case ShtMdy
    Case "Pub", "": EnsMdy = L
    Case "Prv":     EnsMdy = "Private " & L
    Case "Frd":     EnsMdy = "Friend " & L
    Case Else
        Thw CSub, "Given parameter [ShtMdy] must be ['' Pub Prv Frd]", "ShtMdy", ShtMdy
    End Select
End Function
