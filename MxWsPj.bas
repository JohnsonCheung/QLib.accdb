Attribute VB_Name = "MxWsPj"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxWsPj."

Function MdzWs(S As Worksheet) As CodeModule
Set MdzWs = CmpzWs(S).CodeModule
End Function

Function CmpzWs(S As Worksheet) As VBComponent
Set CmpzWs = FstzItn(PjzWs(S).VBComponents, S.CodeName)
End Function

Function PjzWb(b As Workbook) As VBProject
Set PjzWb = b.VBProject
End Function

Function PjzWs(S As Worksheet) As VBProject
Set PjzWs = WbzWs(S).VBProject
End Function

Function PjzRg(A As Range) As VBProject
Set PjzRg = WbzRg(A).VBProject
End Function
