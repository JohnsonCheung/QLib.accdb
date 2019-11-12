Attribute VB_Name = "MxRes"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRes."

Function Resl$(ResFn$, Optional ResPseg$)
Dim F$: F = ResFfn(ResFn, ResPseg)
If NoFfn(F) Then Exit Function
Resl = LineszFt(ResFfn(ResFn, ResPseg))
End Function

Sub WrtRes(S$, ResFn$, Optional ResPseg$, Optional OvrWrt As Boolean)
Dim Ft$: Ft = ResFfn(ResFn, ResPseg)
WrtStr S, Ft, OvrWrt
End Sub

Function Res(ResFn$, Optional ResPseg$) As String()
Res = SplitCrLf(Resl(ResFn, ResPseg))
End Function

Function ResHom$()
ResHom = AddFdrEns(CPjPth, CPjfn & ".res")
End Function

Function ResPth$(ResPseg$)
ResPth = AddPsegEns(ResHom, ResPseg)
End Function

Function ResFfn$(ResFn$, Optional ResPseg$)
'Ret : :Ft #Resource-Ffn#
ResFfn = ResPth(ResPseg) & ResFn
End Function

Function ResDrs(ResFnn$, Optional ResPseg$) As Drs
ResDrs = DrszFcsv(ResFcsv(ResFnn, ResPseg))
End Function
Function ResLo(ResFnn$, Optional Pseg$) As ListObject
Dim F$: F = ResFcsv(ResFnn, Pseg)
OpnFcsv F
Set ResLo = CrtLo(RgzAldta(FstWs(LasWb)))
End Function

Function ResFcsv$(ResFnn$, Optional Pseg$)
ResFcsv = ResFfn(ResFnn & ".csv", Pseg)
End Function

Sub WrtResLoMdP()
WrtDrs DoMdP, ResFcsv("DoMdP")
End Sub

Function ResLoMdP() As ListObject
Set ResLoMdP = ShwLo(ResLo("DoMdP"))
End Function

Sub WrtResLoMthP()
WrtDrs DoMthP, ResFcsv("DoMthP")
End Sub

Function ResLoMthP() As ListObject
Set ResLoMthP = ShwLo(ResLo("DoMthP"))
End Function

Sub Z_ResDrs()
Dim D As Drs: D = ResDrs("DoMthP")
Stop
End Sub
