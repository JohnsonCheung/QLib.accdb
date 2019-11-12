Attribute VB_Name = "MxCurIde"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCurIde."
Function CCmp() As VBComponent
Set CCmp = CMd.Parent
End Function

Function CMd() As CodeModule
Dim P As CodePane: Set P = CPne
If IsNothing(P) Then Exit Function
Set CMd = P.CodeModule
End Function

Function CMdn()
CMdn = CCmp.Name
End Function

Function CLno&()
CLno = CLnozM(CMd)
End Function

Function CMdDn$()
CMdDn = MdDn(CMd)
End Function

Property Get CMthl$() 'Cur
CMthl = MthlzM(CMd, CMthn)
End Property

Function CMthn$()
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Function
CMthn = CMthnzM(CMd)
End Function

Function CWin() As VBIDE.Window
Dim A As CodePane: Set A = CPne
If IsNothing(A) Then Exit Function
Set CWin = A.Window
End Function

Function CPne() As VBIDE.CodePane
Set CPne = CVbe.ActiveCodePane
End Function

Function CMthLno&()
CMthLno = MthLno(CMd, CLno)
End Function

Function CVbe() As Vbe
Set CVbe = Application.Vbe
End Function

Function CPjPth$()
CPjPth = Pth(CPjf)
End Function

Function CPjf$()
CPjf = CPj.FileName
End Function

Function CPjfn$()
CPjfn = Fn(CPj.FileName)
End Function

Function CPj() As VBProject
Set CPj = CVbe.ActiveVBProject
End Function

Function CPjn$()
CPjn = CPj.Name
End Function

Function CMthDn$()
On Error GoTo X
CMthDn = MthDnzM(CMd, CMthLin)
Exit Function
X: Debug.Print CSub
End Function
