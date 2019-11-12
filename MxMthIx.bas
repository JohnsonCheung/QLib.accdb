Attribute VB_Name = "MxMthIx"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthIx."

Function MthEndLin$(Src$(), ItmIx)
MthEndLin = EndLinzVbaItm(VbaItm(Src(ItmIx)))
End Function

Function MthEndLinzM$(M As CodeModule, ItmLno)
MthEndLinzM = EndLinzVbaItm(VbaItm(M.Lines(ItmLno, 1)))
End Function

Function EndLinzVbaItm$(Itm$)
If Not IsVbaItm(Itm) Then Thw CSub, "Given Itm is not a VbaItm", "Itm", Itm
EndLinzVbaItm = "End " & Itm
End Function

Function EndLix&(Src$(), MthIx)
Dim EndL$, O&
EndL = MthEndLin(Src, MthIx)
If HasSubStr(Src(MthIx), EndL) Then EndLix = MthIx: Exit Function
For O = MthIx + 1 To UB(Src)
   If HasPfx(Src(O), EndL) Then EndLix = O: Exit Function
Next
Thw CSub, "Cannot find MthEndLin", "MthEndLin MthIx Src", EndL, MthIx, Src
End Function

Function EndLnozM&(M As CodeModule, ItmLno)
Dim EndL$, O&
EndL = MthEndLinzM(M, ItmLno)
If HasSubStr(M.Lines(ItmLno, 1), EndL) Then EndLnozM = ItmLno: Exit Function
For O = ItmLno + 1 To M.CountOfLines
   If HasPfx(M.Lines(O, 1), EndL) Then EndLnozM = O: Exit Function
Next
Thw CSub, "Cannot find MthEndLin", "MthEndLin ItmLno Md", EndL, ItmLno, Mdn(M)
End Function

Function FstMthIxzN&(Src$(), Mthn, Optional SrcFmIx& = 0)
Dim I
For I = SrcFmIx To UB(Src)
    If MthnzLin(Src(I)) = Mthn Then
        FstMthIxzN = I
        Exit Function
    End If
Next
FstMthIxzN = -1
End Function

Function FstMthIx&(Src$())
Dim O&
    Dim L: For Each L In Itr(Src)
        If IsLinMth(L) Then FstMthIx = O: Exit Function
    Next
FstMthIx = -1
End Function

Function FstMthIxzSN&(Src$(), Mthn)
Dim Ix&
For Ix = 0 To UB(Src)
    With Mthn3zL(Src(Ix))
        If .Nm = Mthn Then
            FstMthIxzSN = Ix
            Exit Function
        End If
    End With
Next
FstMthIxzSN = -1
End Function

Function FstMthLnozM&(Md As CodeModule)
Dim J&
For J = 1 To Md.CountOfLines
   If IsLinMth(Md.Lines(J, 1)) Then
       FstMthLnozM = J
       Exit Function
   End If
Next
End Function

Function IsVbaItm(Itm$) As Boolean
':VbaItm: :S ! One of :VbaItmAy
':VbaItmAy: :Ny ! One of {Function Sub Type Enum Property Dim Const}
IsVbaItm = HasEle(VbaItmAy, Itm)
End Function

Function MthEIxy(Src$(), FmIxy&()) As Long()
Dim Ix, ELin$
Stop
For Each Ix In Itr(FmIxy)
    PushI MthEIxy, EndLix(Src, Ix)
Next
End Function

Function MthFeizSIW(Src$(), MthIx) As Fei
MthFeizSIW = Fei(MthIx, EndLix(Src, MthIx))
End Function

Function MthIxItr(Src$())
Asg Itr(MthIxy(Src)), MthIxItr
End Function

Sub Z_MthIxy()
Dim Src$()
GoSub Z
Exit Sub
Z:
    Src = SrczMdn("QVb_Fs_Pth")
    Dim MIxy&(): MIxy = MthIxy(Src)
    Brw AwIxy(Src, MIxy)
    Return

End Sub

Function MthIxy(Src$()) As Long()
Dim Ix&: For Ix = 0 To UB(Src)
    If IsLinMth(Src(Ix)) Then
        PushI MthIxy, Ix
    End If
Next
End Function

Function MthIxyzM(M As CodeModule, Mthn) As Long()
MthIxyzM = MthIxyzN(Src(M), Mthn)
End Function

Function MthIxyzN(Src$(), Mthn) As Long()
Dim Ix&
Ix = FstMthIxzN(Src, Mthn)
If Ix = -1 Then Exit Function
PushI MthIxyzN, Ix
If IsLinPrp(Src(Ix)) Then
    Ix = FstMthIxzN(Src, Mthn, Ix + 1)
    If Ix > 0 Then PushI MthIxyzN, Ix
End If
End Function

Function MthIxyzSNy(Src$(), MthNy$()) As Long()
Dim Ix: For Each Ix In MthIxItr(Src)
    Dim L$: L = Src(Ix)
    Dim N$: N = Mthn(L)
    If HasEle(MthNy, N) Then PushI MthIxyzSNy, Ix
Next
End Function

Function MthIxzMTN&(M As CodeModule, ShtMthTy$, Mthn)
MthIxzMTN = MthIxzNmTy(Src(M), Mthn, ShtMthTy)
End Function

Function MthIxzNmTy&(Src$(), Mthn, ShtMthTy$)
Dim Ix&
For Ix = 0 To UB(Src)
    With Mthn3zL(Src(Ix))
        If .Nm = Mthn Then
            If .ShtTy = ShtMthTy Then
                MthIxzNmTy = Ix
                Exit Function
            End If
        End If
    End With
Next
MthIxzNmTy = -1
End Function

Function MthLno(Md As CodeModule, Lno&)
Dim O&
For O = Lno To 1 Step -1
    If IsLinMth(Md.Lines(O, 1)) Then MthLno = O: Exit Function
Next
End Function

Function MthLnoAy(M As CodeModule, Mthn) As Long()
MthLnoAy = AmIncEleBy1(MthIxyzN(Src(M), Mthn))
End Function

Function MthLnozMM&(M As CodeModule, Mthn, Optional IsInf As Boolean)
Dim L&: L = FstMthIxzN(Src(M), Mthn, 0)
If L = -1 Then
    If IsInf Then
        Debug.Print "Mth[" & Mthn & "] in Md[" & Mdn(M) & "] not found"
    End If
    Exit Function
End If
MthLnozMM = 1 + L
End Function

Function VbaItm$(Lin)
Dim O$: O = T1(RmvMdy(Lin))
If IsVbaItm(O) Then VbaItm = O
End Function

Function VbaItmAy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = SyzSS("Function Sub Type Enum Property Dim Const Option Implements")
End If
VbaItmAy = Y
End Function

Function VbaItmAyV() As String()
VbaItmAyV = VbaItmAyzSrc(SrcV)
End Function

Function VbaItmAyzSrc(Src$()) As String()
Dim S
For Each S In Itr(Src)
    PushNB VbaItmAyzSrc, VbaItm(S)
Next
End Function
