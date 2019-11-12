Attribute VB_Name = "MxCmpInf"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCmpInf."
Function PjzC(A As VBComponent) As VBProject
Set PjzC = A.Collection.Parent
End Function
Function HasCmp(Cmpn) As Boolean
HasCmp = HasCmpzP(CPj, Cmpn)
End Function
Function HasCmpzP(P As VBProject, Cmpn) As Boolean
If IsProtectzvInf(P) Then Exit Function
HasCmpzP = HasItn(P.VBComponents, Cmpn)
End Function

Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function

Function HasCmpzPTN(P As VBProject, Ty As vbext_ComponentType, Cmpn) As Boolean
Dim T As vbext_ComponentType
If Not HasCmpzP(P, Cmpn) Then Exit Function
T = CmpTyzPN(P, Cmpn)
If T = Ty Then HasCmpzPTN = True: Exit Function
Thw CSub, "Pj has Cmp not as expected type", "PjnzC EptTy ActTy", P.Name, Cmpn, ShtCmpTy(Ty), ShtCmpTy(T)
End Function

Function MdAyzC(CmpAy() As VBComponent) As CodeModule()
Dim I
For Each I In Itr(CmpAy)
    PushObj MdAyzC, CvCmp(I).CodeModule
Next
End Function

Function TmpMod() As CodeModule
Dim T$: T = TmpNm("TmpMod")
AddModnzP CPj, T
Set TmpMod = Md(T)
End Function

Function TmpModNyzP(P As VBProject) As String()
TmpModNyzP = AwPfx(ModNyzP(P), "TmpMod")
End Function
