Attribute VB_Name = "MxTst"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxTst."
Public Act, Ept, Dbg As Boolean, Trc As Boolean
Sub Can_A_AyDic_To_Be_Pushed()
Dim A As Dictionary, Act, V
GoSub T1
Exit Sub
T1: 'This fail
    Set A = New Dictionary
    A.Add "A", EmpAv
    PushI A("A"), 1
    V = A("A")
    Act = Si(V)
    If Si(Act) <> 1 Then Stop
    Return
T2:  'Should Pass
    Set A = New Dictionary
    A.Add "A", EmpAv
    V = A("A")
    PushI V, 1
    A("A") = V
    Act = A("A")
    If Si(Act) <> 1 Then Stop
    Return
'Ans is: Cannot
End Sub
Function TstHom$()
TstHom = TstHomP
End Function
Function TstHomP$()
Static X$: If X = "" Then X = TstHomzP(CPj)
TstHomP = X
End Function
Function TstHomzP$(P As VBProject)
TstHomzP = AddFdrEns(SrcpzP(P), ".TstRes")
End Function

Sub StopNE()
If Not IsEq(Act, Ept) Then Stop
End Sub

Sub C()
ThwIf_NE Act, Ept, "Act", "Ept"
End Sub

Function TstCasPth$(TstId&, Cas$)
TstCasPth = AddFdrEns(TstPth(TstId), "Cas-" & Cas)
End Function

Sub BrwTstPth(TstId&, Optional Cas$)
If Cas = "" Then
    BrwPth TstCasPth(TstId, Cas)
Else
    BrwPth TstPth(TstId)
End If
End Sub

Function TstPth$(TstId&)
If IsNBet(TstId, 0, 9999) Then Thw CSub, "TstId should be 0 to 9999", "TstId", TstId

TstPth = AddFdrEns(TstHom, Pad0(TstId, 4))
End Function

Function TstIdFt$(TstId&)
TstIdFt = TstPth(TstId) & "TstId.Txt"
End Function
Sub BrwTstIdPth(TstId&)
BrwPth TstPth(TstId)
End Sub

Sub BrwTstHom()
BrwPth TstHom
End Sub
Function NxtIdFdr$(Pth, Optional NDig& = 4) '
Dim J&, F$
ThwIf_NoPth Pth, CSub
If NDig < 0 Then Thw CSub, "NDig should between 1 to 5", "NDig", NDig
For J = 1 To Val(Left("99999", NDig))
    F = Pad0(J, NDig)
    If Not HasFdr(Pth, F) Then NxtIdFdr = F: Exit Function
Next
Thw CSub, "IdFdr is full in Pth", "Pth NDig", "Pth NDig", Pth, NDig
End Function
Function NxtTstId%()
NxtTstId = NxtIdFdr(TstHom, 4)
End Function
Sub ShwTstOk(Fun$, Cas$)
Debug.Print "Tst OK | "; Fun; " | Case "; Cas
End Sub

Function TstLy(TstId&, Fun$, Cas$, Itm$, Optional IsEdt As Boolean) As String()
TstLy = SplitCrLf(TstTxt(TstId, Fun, Cas, Itm, IsEdt))
End Function
Function TstIdStr$(TstId&, Fun$)
TstIdStr = "TstId=" & TstId & ";CSub=" & Fun
End Function
Sub WrtTstPth(TstId&, Fun$)
Dim F$: F = TstIdFt(TstId)
Dim IdStr$: IdStr = TstIdStr(TstId, Fun)
Dim Exist As Boolean
Exist = HasFfn(F)
Select Case True
Case (Exist And LineszRTrim(LineszFt(F)) <> IdStr) Or Not Exist
    WrtStr TstIdStr(TstId, Fun), F
End Select
End Sub
Sub EnsTstPth(TstId&, Fun$)
Dim F$: F = TstIdFt(TstId)
Dim IdStr$: IdStr = TstIdStr(TstId, Fun)
Dim Exist As Boolean
Exist = HasFfn(F)
Select Case True
Case Exist And LineszRTrim(LineszFt(F)) <> IdStr
    Thw CSub, "TstIdStr in TstIdFt is not expected", "TstIdFt Expected-TstIdStr Actual-TstIdStr-in-TstIdFt", F, IdStr, LineszFt(F)
Case Exist:
Case Else
    WrtStr TstIdStr(TstId, Fun), F
End Select
End Sub

Function TstTxt$(TstId&, Fun$, Cas$, Itm$, Optional IsEdt As Boolean)
Dim F$:                   F = TstFt(TstId, Cas, Itm)
Dim Exist As Boolean: Exist = HasFfn(F)
                              EnsTstPth TstId, Fun
Select Case True
Case Not Exist: EnsFt F: BrwFt F: Stop
Case IsEdt:     BrwFt F:           Stop
End Select
TstTxt = LineszFt(F)
End Function

Function TstFt$(TstId&, Cas$, Itm$)
TstFt = TstFfn(TstId, Cas, Itm & ".Txt")
End Function

Function TstFfn$(TstId&, Cas$, Fn$)
TstFfn = TstCasPth(TstId, Cas) & Fn
End Function
