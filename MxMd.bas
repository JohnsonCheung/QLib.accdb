Attribute VB_Name = "MxMd"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMd."
':DiMdDnqSrcl: :Dic ! It is from Pj. Key is Mdn and Val is MdLines"
':MdDn: :Pjn.Mdn|Mdn

Function Md(MdDn) As CodeModule
Const CSub$ = CMod & "Md"
Dim A1$(): A1 = Split(MdDn, ".")
Select Case Si(A1)
Case 1: Set Md = CPj.VBComponents(A1(0)).CodeModule
Case 2: Set Md = Pj(A1(0)).VBComponents(A1(1)).CodeModule
Case Else: Thw CSub, "[MdDn] should be XXX.XXX or XXX", "MdDn", MdDn
End Select
End Function

Function MdDn$(M As CodeModule)
MdDn = PjnzM(M) & "." & Mdn(M)
End Function

Function DiMdnqSrclzP(P As VBProject) As Dictionary
Dim C As VBComponent
Set DiMdnqSrclzP = New Dictionary
For Each C In P.VBComponents
    DiMdnqSrclzP.Add C.Name, Srcl(C.CodeModule)
Next
End Function

Function DiMdnqSrclP() As Dictionary
Set DiMdnqSrclP = DiMdnqSrclzP(CPj)
End Function

Function MdFn$(M As CodeModule)
MdFn = Mdn(M) & ExtzCmpTy(CmpTyzM(M))
End Function

Function Mdn(M As CodeModule)
Mdn = M.Name
End Function

Function MdnzM(M As CodeModule)
MdnzM = M.Parent.Name
End Function

Function MdTy(M As CodeModule) As vbext_ComponentType
MdTy = M.Parent.Type
End Function

Function ShtCmpTyzM$(M As CodeModule)
ShtCmpTyzM = ShtCmpTy(CmpTyzM(M))
End Function

Function PjnzC(A As VBComponent)
PjnzC = A.Collection.Parent.Name
End Function

Function PjnzM(M As CodeModule)
PjnzM = PjnzC(M.Parent)
End Function

Function PjzM(M As CodeModule) As VBProject
Set PjzM = M.Parent.Collection.Parent
End Function

Function SizMd&(M As CodeModule)
SizMd = Len(Srcl(M))
End Function

Function RmvMthInSrc(Src$(), MthnSet As Aset) As String()
Dim D As Dictionary: 'Set D = DiceKeySet(MthnDic(Src), MthnSet): 'Brw D: Stop
RmvMthInSrc = LyzLinesDicItems(D)
End Function


Function HasBar(BarNm) As Boolean
HasBar = HasBarzV(CVbe, BarNm)
End Function

Function HasPjf(Pjf) As Boolean
HasPjf = HasPjfzV(CVbe, Pjf)
End Function

Function PjzPjfC(Pjf) As VBProject
Set PjzPjfC = PjzPjf(CVbe, Pjf)
End Function

Function HasMd(P As VBProject, Mdn, Optional IsInf As Boolean) As Boolean
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Name = Mdn Then HasMd = True: Exit Function
Next
If IsInf Then
    Debug.Print FmtQQ("Mdn[?] not exist", Mdn)
End If
End Function

Sub ThwIf_NotMod(M As CodeModule, Fun$)
If Not IsMod(M) Then Thw Fun, "Should be a Mod", "Mdn MdTy", Mdn(M), ShtCmpTy(CmpTyzM(M))
End Sub

Function HasMod(P As VBProject, Modn) As Boolean
If Not HasMd(P, Modn) Then Exit Function
ThwIf_NotMod MdzP(P, Modn), CSub
End Function

Function PjnyzX(X As Excel.Application) As String()
PjnyzX = PjNyzV(X.Vbe)
End Function

Property Get PjnyX() As String()
PjnyX = PjnyzX(Xls)
End Property

Sub SavCurVbe()
SavVbe CVbe
End Sub

Sub ClsMd(M As CodeModule)
M.CodePane.Window.Close
End Sub

Sub CprMd(M As CodeModule, B As CodeModule)
'BrwCprDic DiMthnqLineszM(A), DiMthnqLineszMd(B), MdDn(A), MdDn(B)
End Sub

Function Y_Md() As CodeModule
Set Y_Md = CVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Function

Sub Z_DroMds()
'BrwDrs DroMds(Md("IdeFeature_EnsZ_AsPrivate"))
End Sub

Sub Z_MthLnozMM()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MthNyzM(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, MthLnozMM(A, CStr(M))
        If J Mod 150 = 0 Then
            Debug.Print J, Si(Ny), "Z_MthLnozMM"
        End If
    Next

    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
BrwAy O
End Sub

Sub Z_CMd()
Ass CMd.Parent.Name = "Cur_d"
End Sub

Function CvMd(A) As CodeModule
Set CvMd = A
End Function

Function LibNyP() As String()
LibNyP = LibNyzP(CPj)
End Function

Function LibNyzP(P As VBProject) As String()
LibNyzP = AeBlnk(AwDistAsSy(AmBef(MdNyzP(P), "_")))
End Function
