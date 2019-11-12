Attribute VB_Name = "MxGenPj"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxGenPj."
':SrcRoot: :Pth #Src-Root#          ! is a :Pth.  Its :Fdr eq ".Src"
':Srcp:    :Pth #Src-Path#          ! is a :Pth.  Its :Fdr is a PjFn and par fdr is :SrcRoot
':Disp:    :Pth #Distribution-Path# ! is a :Pth.  d
':InstPth: :Pth #Instance-Path#     ! of a @pth is any :TimNm :Fdr under @pth
':TimNm:   :Nm

Function Fxa$(FxaNm, Srcp)
Fxa = Distp(Srcp) & FxaNm & ".xlam"
End Function

Function Fba$(FbaNm, Srcp)
Fba = EnsPth(Srcp & "Dist") & FbaNm & ".accdb"
End Function

Sub Z_CompressFxa()
CompressFxa Pjf(CPj)
End Sub

Sub CompressFxa(Fxa$)
ExpPjzP PjzPjf(Xls.Vbe, Fxa)
Dim Srcp$: Srcp = SrcpzPjf(Fxa)
GenFxazSrcp Srcp
'BackupFfn Fxa, Srcp
End Sub

Function SrcRoot$(Srcp$)
'Ret: :SrcRoot @@
SrcRoot = ParPth(Srcp)
End Function

Function DistpP$() 'Distribution Path
DistpP = Distp(SrcpP)
End Function

Function Distp$(Srcp)
Dim A$: A = DltPthSfx(Srcp)
If HasPfx(A, ".Src") Then Thw CSub, "Given @Srcp is not a Src path (fdr with .src at end)", "Given-Src-Pth", Srcp
Distp = EnsPth(RmvExt(A) & ".dist")
End Function

Function DistFbaP$()
DistFbaP = DistFba(SrcpzP(CPj))
End Function

Function DistFba$(Srcp)
DistFba = DistPjfzSrcp(Srcp, ".accdb")
End Function

Function DistPjfzSrcp(Srcp, Ext)
Dim P$:   P = Distp(Srcp)
Dim F1$: F1 = RplExt(Fdr(ParPth(P)), Ext)
Dim F2$: F2 = NxtFfnzNotIn(F1, PjfnAyV)
Dim F$:   F = NxtFfnzAva(P & F2)
DistPjfzSrcp = F
End Function

Sub Z_FxazSrcp()
Dim Srcp$
GoSub T0
Exit Sub
T0:
    Srcp = SrcpP
    Ept = "C:\Users\user\Documents\Projects\Vba\QLib\.Dist\QLib(002).xlam"
    GoTo Tst
Tst:
    Act = DistFxazSrcp(Srcp)
    C
    Return
End Sub

Function DistFxazSrcp$(Srcp)
DistFxazSrcp = DistPjfzSrcp(Srcp, ".xlam")
End Function

Sub LoadBas(P As VBProject, Srcp$)
Dim F$(): F = BasFfnAy(Srcp)
Dim I: For Each I In Itr(F)
    P.VBComponents.Import I
Next
End Sub
Sub LoadBas3(P As VBProject, Srcp$)
Dim F$(): F = BasFfnAy(Srcp)
Dim J%, I: For Each I In Itr(F)
    P.VBComponents.Import I
    J = J + 1
    If J > 3 Then Exit Sub
Next
End Sub

Function BasFfnAy(Srcp$) As String()
Dim F$(): F = FfnAy(Srcp)
Dim I: For Each I In Itr(F)
    If IsBasFfn(I) Then
        PushI BasFfnAy, I
    End If
Next
End Function

Function IsBasFfn(Ffn) As Boolean
IsBasFfn = HasSfx(Ffn, ".bas")
End Function

Sub GenFbaP()
GenFbazP CPj
End Sub

Sub GenFbazP(P As VBProject)
Dim OPj As VBProject
Dim SPth$:     SPth = SrcpzP(P)         '#Src-Pth#

Dim OFba$:     OFba = DistFba(SPth)     '#Oup-Fba#
:                     DltFfnIf OFba
:                     CrtFb OFba                    ' <== Crt OFba

:                     ExpPjzP P                       ' <== Exp

:                     OpnFb Acs, OFba
            Set OPj = PjzAcs(Acs)
:                     AddRfzS OPj, RfSrczSrcp(SPth) ' <== Add Rf
:                     LoadBas OPj, SPth             ' <== Load Bas
Dim Frm$():     Frm = FrmFfnAy(SPth)
Dim F: For Each F In Itr(Frm)
    Dim N$: N = RmvExt(RmvExt(F))
:               Acs.LoadFromText acForm, N, F       ' <== Load Frm
Next
#If False Then
'Following code is not able to save
Dim Vbe As Vbe: Set Vbe = Acs.Vbe
Dim C As VBComponent: For Each C In Acs.Vbe.ActiveVBProject.VBComponents
    C.Activate
    BoSavzV(Vbe).Execute
    Acs.Eval "DoEvents"
Next
#End If
MsgBox "Go access to save....."
Inf CSub, "Fba is created", "Fba", OFba
End Sub

Sub GenFxaP()
GenFxazP CPj
End Sub
Sub GenFxazSrcp(Srcp$)

End Sub

Sub GenFxazP(Pj As VBProject)
Dim SPth$:               SPth = SrcpzP(Pj)
Dim OFxa$:               OFxa = DistFxazSrcp(SPth)
:                               ExpPjzP Pj                                 ' <== Export
:                               CrtFxa OFxa                              ' <== Crt
Dim OPj As VBProject: Set OPj = PjzFxa(OFxa)
:                               AddRfzS OPj, RfSrczSrcp(SPth)            ' <== Add Rf
:                               LoadBas OPj, SPth                        ' <== Load Bas
:                               Inf CSub, "Fxa is created", "Fxa", OFxa
End Sub

