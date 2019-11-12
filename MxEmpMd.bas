Attribute VB_Name = "MxEmpMd"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEmpMd."
Function IsMdEmp(M As CodeModule) As Boolean
If M.CountOfLines = 0 Then IsMdEmp = True: Exit Function
Dim J&, L$
For J = 1 To M.CountOfLines
    If Not IsLinEmpSrc(M.Lines(J, 1)) Then Exit Function
Next
IsMdEmp = True
End Function

Sub RmvEmpMd()
Dim N
For Each N In Itr(EmpMdNy)
    RmvCmp CPj.VBComponents(N)
Next
End Sub

Property Get EmpMdNy() As String()
EmpMdNy = EmpMdNyzP(CPj)
End Property

Function EmpMdNyzP(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsMd(C) Then
        If IsMdEmp(C.CodeModule) Then
            PushI EmpMdNyzP, C.Name
        End If
    End If
Next
End Function

Sub Z_IsMdEmp()
Dim M As CodeModule
'GoSub T1
'GoSub T2
GoSub T3
Exit Sub
T3:
    Debug.Assert IsMdEmp(Md("Dic"))
    Return
T2:
    Set M = Md("Module2")
    Ept = True
    GoTo Tst
T1:
    '
    Dim T$, P As VBProject
        Set P = CPj
        T = TmpNm
    '
'    Set M = PjAddMd(P, T)
    Ept = True
    GoSub Tst
    DltCmpzPjn P, T
    Return
Tst:
    Act = IsMdEmp(M)
    C
    Return
End Sub

Function IsSrcEmp(A$()) As Boolean
Dim L
For Each L In Itr(A)
    If Not IsLinEmpSrc(L) Then Exit Function
Next
IsSrcEmp = True
End Function

Function EmpMdNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy EmpMdNyzV, EmpMdNyzP(P)
Next
End Function

Function IsMdNoMth(M As CodeModule) As Boolean
Dim J&
For J = M.CountOfDeclarationLines + 1 To M.CountOfLines
    If IsLinMth(M.Lines(J, 1)) Then Exit Function
Next
IsMdNoMth = True
End Function
