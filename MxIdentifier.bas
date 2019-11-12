Attribute VB_Name = "MxIdentifier"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdentifier."
Sub Z_NyzStr()
Dim S$
GoSub Z
'GoSub T0
Exit Sub
Z:
    Dim Lines$: Lines = SrclP
    Dim Ny1$(): Ny1 = NyzStr(Lines)
    Dim Ny2$(): Ny2 = WrdAy(Lines)
    If Not IsEqAy(Ny1, Ny2) Then Stop
    Return
T0:
    S = "S_S"
    Ept = Sy("S_S")
    GoTo Tst
Tst:
    Act = NyzStr(S)
    C
    Return
End Sub
Sub Z_NsetzStr()
NsetzStr(SrclP).Srt.Vc
End Sub
Function NsetzStr(S) As Aset
Set NsetzStr = AsetzAy(NyzStr(S))
End Function

Function NyzStr(S) As String()
NyzStr = AwNoNm(SyzSS(RplLf(RplCr(RplPun(S)))))
End Function

Function RelOfPubMthn_ToMdny_P() As Rel
Set RelOfPubMthn_ToMdny_P = RelOfPubMthn_ToMdny_zP(CPj)
End Function
Function RelOfMthn_ToCml_V() As Rel
Set RelOfMthn_ToCml_V = RelOfMthn_ToCml_zV(CVbe)
End Function

Function RelOfMthn_ToCml_zV(A As Vbe) As Rel
Dim O As New Rel, I
For Each I In MthNyzV(A)
    O.PushRelLin Cmlss(I)
Next
Set RelOfMthn_ToCml_zV = O
End Function

Function RelOfPubMthn_ToMdny_zP(P As VBProject) As Rel
Dim C As VBComponent, S$(), O As New Rel, Mthn, Modn, Cmp As VBComponent
For Each C In P.VBComponents
    Set Cmp = C
    Modn = Cmp.Name
    S = Src(Cmp.CodeModule)
    For Each Mthn In Itr(MthNy(S))
        O.PushParChd Mthn, C.Name
    Next
Next
Set RelOfPubMthn_ToMdny_zP = O
End Function
Function RelOfMthn_ToMdny_zP(P As VBProject) As Rel
Dim C As VBComponent, O As New Rel, Mthn, Mdn
For Each C In P.VBComponents
    Mdn = C.Name
    For Each Mthn In Itr(MthNy(Src(C.CodeModule)))
        O.PushParChd Mthn, Mdn
    Next
Next
Set RelOfMthn_ToMdny_zP = O
End Function
Function MthnzRlMdnP() As Rel
Static O As Rel
If IsNothing(O) Then Set O = RelOfMthn_ToMdny_zP(CPj)
Set MthnzRlMdnP = O
End Function
Function MthExtny(MthPjDotMdn, PubMthLy$(), PubMthn_To_PjDotModNy As Dictionary) As String()
Dim Cxt$: Cxt = JnSpc(MthCxtLy(PubMthLy))
Dim Ny$(): Ny = NyzStr(Cxt)
Dim Nm
For Each Nm In Itr(Ny)
    If PubMthn_To_PjDotModNy.Exists(Nm) Then
        Dim PjDotModNy$():
            PjDotModNy = AeEle(PubMthn_To_PjDotModNy(Nm), MthPjDotMdn)
        If HasEle(PjDotModNy, Nm) Then
            PushI MthExtny, Nm
        End If
    End If
Next
End Function

Property Get VbKwAy() As String()
Static X$()
If Si(X) = 0 Then
    X = SyzSS("Function Sub Then If As For To Each End While Wend Loop Do Static Dim Option Explicit Compare Text")
End If
VbKwAy = X
End Property

Property Get VbKwAset() As Aset
Set VbKwAset = AsetzAy(VbKwAy)
End Property
