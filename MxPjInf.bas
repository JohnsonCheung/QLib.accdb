Attribute VB_Name = "MxPjInf"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxPjInf."

Function PjzFxa(Fxa) As VBProject
'Ret: Ret :Pj of @Fxa fm @Xls if exist, else @Xls.Opn @Fxa
Dim O As VBProject: Set O = PjzPjf(Xls.Vbe, Fxa)
If Not IsNothing(O) Then Set PjzFxa = O: Exit Function
Set PjzFxa = OpnFx(Fxa).VBProject
End Function

Function HasFxa(Fxa$) As Boolean
HasFxa = HasEleS(PjfnAyV, Fn(Fxa))
End Function

Sub OpnFxa(Fxa$)
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then
    Inf CSub, "In Xls, there is Pjn = Fxa", "Fxa AllPj-In-Xls", Fxa, PjNyV
    Exit Sub
End If
Xls.Workbooks.Open Fxa
End Sub

Function PjnzFxa$(Fxa)
PjnzFxa = Fnn(RmvNxtNo(Fxa))
End Function

Sub CrtFxa(Fxa$)
'Do: crt an emp Fxa with pjn derived from @Fxa
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then Thw CSub, "In Xls, there is Pjn = Fxa", "Fxa AllPj-In-Xls", Fxa, PjNyV
Dim Wb As Workbook: Set Wb = Xls.Workbooks.Add
Wb.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn 'Must save first, otherwise PjzFxa will fail.
PjzFxa(Fxa).Name = PjnzFxa(Fxa)
Wb.Close True
End Sub

Function FrmFfnAy(Pth) As String()
Dim I: For Each I In Itr(FfnAy(Pth, "*.frm.txt"))
    PushI FrmFfnAy, I
Next
End Function

Function ClsAyzP(P As VBProject) As CodeModule()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsCls(C) Then
        PushObj ClsAyzP, C
    End If
Next
End Function

Function ClsNyzP(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsCls(C) Then
        PushI ClsNyzP, C.Name
    End If
Next
End Function

Sub Z_CmpAyzP()
Dim Act() As VBComponent
Dim C, T As vbext_ComponentType
For Each C In CmpAyzP(CPj)
    T = CvCmp(C).Type
    If T <> vbext_ct_StdModule And T <> vbext_ct_ClassModule Then Stop
Next
End Sub

Function CmpAyzP(P As VBProject) As VBComponent()
If IsProtectzvInf(P) Then Exit Function
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMd(C) Then
        PushObj CmpAyzP, C
    End If
Next
End Function

Function IsPjNoClsNoMod(P As VBProject) As Boolean
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
IsPjNoClsNoMod = True
End Function

Function ModItrzP(P As VBProject)
Asg Itr(ModAyzP(P)), ModItrzP
End Function

Function ModAyzP(P As VBProject) As CodeModule()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent: For Each C In P.VBComponents
    If C.Type = vbext_ct_StdModule Then
        PushObj ModAyzP, C.CodeModule
    End If
Next
End Function

Function MdNyP() As String()
MdNyP = MdNyzP(CPj)
End Function

Function MdNyoNoLin() As String()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If C.CodeModule.CountOfLines = 0 Then
        PushI MdNyoNoLin, C.Name
    End If
Next
End Function

Function MdNyWiPrpV() As String()
MdNyWiPrpV = MdNyWiPrpzV(CVbe)
End Function

Function MdNyWiPrpzV(A As Vbe) As String()
Dim Mdn, I
For Each I In MdNyzV(A)
    Mdn = I
    If IsMdnWiPrp(Mdn) Then
        PushI MdNyWiPrpzV, Mdn
    End If
Next
End Function

Function IsMdnWiPrp(Mdn) As Boolean
Dim M As CodeModule: Set M = Md(Mdn)
Dim J&
For J = 1 To M.CountOfLines
    If IsLinPrp(M.Lines(J, 1)) Then IsMdnWiPrp = True: Exit Function
Next
End Function

Function MdNyV() As String()
MdNyV = MdNyzV(CVbe)
End Function

Function MdNyzM(Mthn) As String()
MdNyzM = MdnsetzM(Mthn).Sy
End Function

Function MdAyzNN(Mdnn$) As CodeModule()
End Function

Function MdAyzNy(MdNy$()) As CodeModule()
Dim N, P As VBProject
For Each N In Itr(MdNy)
    PushI MdAyzNy, MdzP(P, N)
Next
End Function

Function MdAyzPm(PubMth) As CodeModule()
MdAyzPm = MdAyzNy(MdNyzPm(PubMth))
End Function

Function MdNyzPPm(P As VBProject, PubMthn) As String()
End Function

Function MdNyzPm(PubMthn) As String()
MdNyzPm = MdnsetzPm(PubMthn).Sy
End Function

Function MdnsetzPm(PubMthn) As Aset
Set MdnsetzPm = PubMthnzRlMdnV.ParChd(PubMthn)
End Function

Function MdnsetzM(Mthn) As Aset
Set MdnsetzM = MthnzRlMdnP.ParChd(Mthn)
End Function

Property Get PubMthnzRlMdnV() As Rel
Set PubMthnzRlMdnV = PubMthnzRlMdnzV(CVbe)
End Property

Function PubMthnzRlMdnzV(A As Vbe) As Rel
Stop
End Function

Function MdNyzP(P As VBProject) As String()
Dim C
For Each C In CmpItr(P)
    If IsMd(CvCmp(C)) Then
        PushI MdNyzP, C.Name
    End If
Next
End Function

Function MdNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MdNyzV, MdNyzP(P)
Next
End Function

Function ModNyP() As String()
ModNyP = ModNyzP(CPj)
End Function

Function ClsNyP() As String()
ClsNyP = ClsNyzP(CPj)
End Function

Function ModNyzP(P As VBProject) As String()
Dim C As VBComponent, O$()
For Each C In P.VBComponents
    If IsMod(C) Then PushI ModNyzP, C.Name
Next
End Function

Sub Z_ClsNyzP()
DmpAy ClsNyzP(CPj)
End Sub

Sub Z_MdAy()
Dim O() As CodeModule
O = MdAyzP(CPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print Mdn(Md)
Next
End Sub

Sub Z_MdzPjny()
'DmpAy MdzPjny(CPj)
End Sub

Function CmpAyP() As VBComponent()
CmpAyP = CmpAyzP(CPj)
End Function

Function MdAy() As CodeModule()
MdAy = MdAyzP(CPj)
End Function

Function CmpItr(P As VBProject)
Asg Itr(CmpAyzP(P)), CmpItr
End Function

Function MdItr(P As VBProject)
Asg Itr(MdAyzP(P)), MdItr
End Function

Function MdAyzP(P As VBProject) As CodeModule()
MdAyzP = MdAyzC(CmpAyzP(P))
End Function


Function SizP&(P As VBProject)
Dim O&, C As VBComponent
For Each C In P.VBComponents
    O = O + SizMd(C.CodeModule)
Next
SizP = O
End Function

Function SiP&()
SiP = SizP(CPj)
End Function

Function DftPj(P As VBProject) As VBProject
If IsNothing(P) Then
    Set DftPj = CPj
Else
    Set DftPj = P
End If
End Function

