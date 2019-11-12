Attribute VB_Name = "MxNmlMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxNmlMd."
Function NmlSrc(XSrc$()) As String() ' #Normaliz-XSrc#
'Ret: Normalized-Src from Exported-Src @@
':XSrc: :Src #Exported-Source# ! after :Cmp.Export, the file with have serval lines added.  This :Src is known as :XSrc
NmlSrc = RmvAtrLines(Rmv4ClassLines(XSrc))
End Function
Function RmvAtrLines(Src$()) As String()
Dim Fm%
    Dim J%: For J = 0 To UB(Src)
        If Not HasPfx(Src(J), "Attribute ") Then
            Fm = J
            GoTo X
        End If
    Next
X:
RmvAtrLines = AwFm(Src, Fm)
End Function

Function Rmv4ClassLines(XSrc$()) As String()
If Si(XSrc) = 0 Then Exit Function
If XSrc(0) = "VERSION 1.0 CLASS" Then
    Rmv4ClassLines = AwFm(XSrc, 4)
Else
    Rmv4ClassLines = XSrc
End If
End Function
