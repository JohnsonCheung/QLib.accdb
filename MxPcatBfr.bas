Attribute VB_Name = "MxPcatBfr"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPcatBfr."
Dim A() As Pcat '<== PcatBfr
Dim N%
Sub Push_Pcat_ToPcatBfr(I As Pcat)
ReDim Preserve A(N)
A(N) = I
N = N + 1
End Sub
Function F_Pcat_ByTar_FmPcatBfr(Tar As Range, OFnd As Boolean) As Pcat
Dim Ws As Worksheet: Set Ws = WszRg(Tar)
Dim RRCC As RRCC: RRCC = RRCCzRg(Tar)
Dim Ix&: Ix = FndIx(Ws, RCzRg(Tar))
If Ix = -1 Then OFnd = False: Exit Function
OFnd = True
F_Pcat_ByTar_FmPcatBfr = A(Ix)
End Function

Function FndIx%(Ws As Worksheet, Tar As RC)
Dim J%: For J = 0 To N - 1
    If IsEqObj(A(J).Ws, Ws) Then
        If HasRC(A(J).UKeyRRCC, Tar) Then
            FndIx = J
            Exit Function
        End If
    End If
Next
FndIx = -1
End Function
