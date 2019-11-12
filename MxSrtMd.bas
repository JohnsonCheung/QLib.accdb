Attribute VB_Name = "MxSrtMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrtMd."
Sub Z_SrtMd()
Dim Md As CodeModule
GoSub X0
Exit Sub
X0:
    Dim I
'    For Each I In MdAy(CPj)
        Set Md = I
        If Mdn(Md) = "Str_" Then
            GoSub Ass
        End If
'    Next
    Return
X1:
    Return
Ass:
    Debug.Print Mdn(Md); vbTab;
    Dim BefSrt$(), AftSrt$()
    BefSrt = Src(Md)
    AftSrt = SplitCrLf(SrtdSrclzM(Md))
    If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then
        Debug.Print "Is Same of before and after sorting ......"
        Return
    End If
    If Si(AftSrt) <> 0 Then
        If LasEle(AftSrt) = "" Then
            Dim Pfx
            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & Mdn(Md) & "=====")
            BrwAy AddAyAp(Pfx, AftSrt)
            Stop
        End If
    End If
    Dim A$(), B$(), II
    A = AyMinus(BefSrt, AftSrt)
    B = AyMinus(AftSrt, BefSrt)
    Debug.Print
    If Si(A) = 0 And Si(B) = 0 Then Return
    If Si(AeEmpEle(A)) <> 0 Then
        Debug.Print "Si(A)=" & Si(A)
        BrwAy A
        Stop
    End If
    If Si(AeEmpEle(B)) <> 0 Then
        Debug.Print "Si(B)=" & Si(B)
        BrwAy B
        Stop
    End If
    Return
End Sub

Sub SrtMdzSrcp(SrcPth)
Dim P$: P = EnsPthSfx(SrcPth)
Dim F: For Each F In FfnAy(P, "*.bas")
    EnsFt F, SrtdSrcl(LyzFt(F))
Next
End Sub

Sub SrtMdzSrcpP()
SrtMdzSrcp SrcpP
End Sub

Sub SrtMd(M As CodeModule)
RplMd M, SrtdSrclzM(M)
End Sub

Function SrtdSrcl$(Src$())
':SrtdSrcL :SrcL #Sorted-Srcl#
SrtdSrcl = JnStrDic(SrtDic(DiMthnqLines(Src)), vb2CrLf)
End Function

Function SrtdSrclM$()
SrtdSrclM = SrtdSrclzM(CMd)
End Function

Function SrtdSrclzM$(M As CodeModule)
SrtdSrclzM = SrtdSrcl(Src(M))
End Function

Function Srcl$(M As CodeModule)
':Srcl: :Lines #Src-Lines#
If M.CountOfLines > 0 Then
    Srcl = M.Lines(1, M.CountOfLines)
End If
End Function

Sub SrtMdM()
SrtMd CMd
End Sub

Function DiMdnqSrtdSrcl(P As VBProject) As Dictionary
Set DiMdnqSrtdSrcl = New Dictionary
Dim C As VBComponent: For Each C In P.VBComponents
    DiMdnqSrtdSrcl.Add C.Name, SrtdSrcl(Src(C.CodeModule))
Next
End Function
