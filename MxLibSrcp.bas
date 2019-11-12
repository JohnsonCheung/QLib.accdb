Attribute VB_Name = "MxLibSrcp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxLibSrcp."

Function LibSrcpzP$(P As VBProject, Libv$)
':LibSrcp: :Pth #Library-Src-Pth# ! a :Srcp under :Libp.
'                                 ! under this :Libp, there are 2-or-more Srcp with Fdr-name = `{Libv}.{Ext}.src`, where {Ext} is Ext-of-Pjf-of-@P.
'                                 ! Note: The folder of the Srcp is in format of `{PjFn}.src`
'                                 ! Example: Given Pjf-@P            "C:\users\user\Documents\Projects\Vba\QLib\QLib.accdb"
'                                 !          Given @Libv             "QVb"
'                                 !          Then LibpzP(@P) will be "C:\users\user\Documents\Projects\Vba\QLib\QLib.accdb.lib\QVb.accdb.src"
Dim Libp$: Libp = LibpzP(P)
Dim Libf$: Libf = 1
LibSrcpzP = EnsPthAll(Libp & Libf)
End Function

Sub EnsLibSrcp(P As VBProject, Libv$)
EnsPthAll LibSrcpzP(P, Libv)
End Sub

Function LibSrcpzDistPj$(DistPj As VBProject)
Dim P$: P = Pjp(DistPj)
LibSrcpzDistPj = AddFdrAp(UpPth(P, 1), ".Src", Fdr(P))
End Function

Function LibSrcpP$(Libv$)
LibSrcpP = LibSrcpzP(CPj, Libv)
End Function

Function LibpzP$(P As VBProject)
':Libp: :Pth #Library-Pth# ! a pj can generate 1 pj or 2-or-more-pj.  When gen 2-or-more-pj, there is a :Libp in the same fdr as the Pjf of @P.
'                          ! under this :Libp, there are 2-or-more Srcp with Fdr-name = `{Libv}.{Ext}.src`, where {Ext} is Ext-of-Pjf-of-@P.
'                          ! Note: The folder of the Srcp is in format of `{PjFn}.src`
'                          ! Example: Given Pjf-@P            "C:\users\user\Documents\Projects\Vba\QLib\QLib.accdb"
'                          !          Then LibpzP(@P) will be "C:\users\user\Documents\Projects\Vba\QLib\QLib.accdb.lib\"
LibpzP = EnsPth(Pjf(P) & ".lib")
End Function

Sub BrwSrcpP()
BrwPth SrcpP
End Sub

Function IsLibSrcp(Pth) As Boolean
Dim F$: F = Fdr(Pth)
If Not HasExtss(F, ".xlam .accdb") Then Exit Function
IsLibSrcp = Fdr(ParPth(Pth)) = ".Src"
End Function
