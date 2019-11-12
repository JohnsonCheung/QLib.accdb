Attribute VB_Name = "MxVerb"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVerb."
Public Const Verbss$ = "Zip Wrt Wrp Wait Vis Vc ULin UnRmk UnEsc Trim Tile Thw Tak Sye Swap Sum Stop Srt Split Solve Shw Shf Set Sel Sav Run Rpl Rmv Rmk Rfh AyRev Resi Ren ReSz ReSeq ReOrd RTrim Qte Quit Push Prompt Pop Opn Nxt Norm New Mov Mk Minus Min Mid Mge Max Map Lnk Lis Lik Las Kill Jn Jmp Is Into AyIntersect Ins Initialize Init Inf Indent Inc Imp Hit Has Halt Gen Fst Fmt Flat Fill Extend Expand Exp Xls Evl Esc Ens EndTrim Edt Dyw Dye Dw De Drp Down Do Dmp Dlt Cv Cut Crt Cpy Compress Cls Clr Clone Cln Clear Chk3 Chk2 Chk1 Chk Chg Change Cfm Brw Brk Box Bld Bet Below Bef Bdr Backup Aw Ae AutoFit AutoExec Ass Asg And Align Aft Add Above"
Public Const C_BRKCmlss$ = "Wi Wo By Of To"
':QBNm$ = "Qte-Brk-Nm.  If the Cml is BRKCml, quote-bkt."
':Rx: :MthSfx #Regular-Expression#
Property Get BRKCmlASet() As Aset
Static X As Aset
If IsNothing(X) Then Set X = AsetzSsl(C_BRKCmlss)
Set BRKCmlASet = X
End Property
Property Get MthVNyInVbe() As String()
Dim Mthn, I
For Each I In Itr(MthNyV)
    Mthn = I
    PushI MthVNyInVbe, MthVNm(Mthn)
Next
End Property
Sub Z_MthVNsetInVbe()
MthVNsetInVbe.Srt.Vc
End Sub
Property Get MthVNsetInVbe() As Aset
Set MthVNsetInVbe = AsetzAy(MthVNyInVbe)
End Property

Function MthQVNsetInVbe() As Aset
Dim Ay$(): Ay = MthQVNyInVbe
Set MthQVNsetInVbe = AsetzAy(Ay)
End Function

Sub VcMthQVNsetInVbe()
MthQVNsetInVbe.Srt.Vc
End Sub

Sub VcMthQVNyInVbe()
AsetzAy(MthQVNyInVbe).Srt.Vc
End Sub

Function MthQVNyInVbe() As String() '6204
MthQVNyInVbe = MthQVNyzV(CVbe)
End Function

Function MthQVNyzV(A As Vbe) As String()
MthQVNyzV = QVNy(MthNyzV(A))
End Function

Function QVNy(Ny$()) As String()
Dim Nm$, I
For Each I In Itr(Ny)
    Nm = I
    PushI QVNy, QVNm(Nm)
Next
End Function

Function QBNm$(Nm)
Dim Cml$, I, O$()
For Each I In Itr(CmlAy(Nm))
    Cml = I
    If IsCmlBRK(Cml) Then
        PushI O, QteBkt(Cml)
    Else
        PushI O, Cml
    End If
Next
QBNm = Jn(O)
End Function

Function QVBNm$(Nm) 'Qte-Verb-and-cmlBrk-Nm.
Dim V$: V = Verb(Nm)
If V = "" Then
    QVBNm = "#" & QBNm(Nm)
Else
    With Brk(Nm, V)
    QVBNm = QBNm(.S1) & QteSq(V) & QBNm(.S2)
    End With
End If
End Function
Function QVNm$(Nm)
Dim V$: V = Verb(Nm)
If V = "" Then
    QVNm = "#" & Nm
Else
    QVNm = Replace(Nm, V, QteBkt(V), Count:=1)
End If
End Function
Function MthVNm$(Mthn)
MthVNm = Verb(Mthn) & "." & Mthn
End Function
Function VerbRx() As RegExp
Static X As RegExp
If IsNothing(X) Then Set X = Rx(VerbPatn(Verbss))
Set VerbRx = X
End Function

Sub BrwVerb()
Brw SyzSS(Verbss)
End Sub

Sub VcNVTDNmAsetInVbe()
NVTDNmAsetInVbe.Srt.Vc
End Sub
Property Get NVTDNmAsetInVbe() As Aset
Set NVTDNmAsetInVbe = AsetzAy(NVTDNyInVbe)
End Property
Property Get NVTDNyInVbe() As String()
NVTDNyInVbe = NVTDNyzV(CVbe)
End Property
Function NVTDNyzV(A As Vbe) As String()
NVTDNyzV = NVTDNy(MthNyzV(A))
End Function
Function NVTDNy(Ny$()) As String()
Dim Nm$, I
For Each I In Itr(Ny)
    Nm = I
    PushI NVTDNy, NVTDNm(Nm)
Next
End Function
Function NVTDNm$(Nm) 'Nm.Verb.Ty.Dot-Nm
NVTDNm = NVTy(Nm) & "." & Nm
End Function
Function FstVerbSubNyInVbe() As String()

End Function
Function NVTy$(Nm) 'Nm.Verb-Ty
Select Case True
Case IsNoVerbNm(Nm): NVTy = "NoVerb"
Case IsFstVerbNm(Nm): NVTy = "FstVerb"
Case IsMidVerbNm(Nm): NVTy = "MidVerb"
Case Else: Thw CSub, "Program error: a Nm must be any of [NoVerb | FstVerb | MidVerb]", "Nm", Nm
End Select
End Function
Function IsNoVerbNm(Nm) As Boolean
IsNoVerbNm = Verb(Nm) = ""
End Function
Function IsMidVerbNm(Nm) As Boolean
Dim V$: V = Verb(Nm): If V = "" Then Exit Function
IsMidVerbNm = Not HasPfx(Nm, Verb(Nm))
End Function

Function IsFstVerbNm(Nm) As Boolean
IsFstVerbNm = HasPfx(Nm, Verb(Nm))
End Function
Function IsVerb(S) As Boolean
IsVerb = VerbAset.Has(RmvEndDig(S))
End Function

Property Get VerbAset() As Aset
Static X As Aset
If IsNothing(X) Then Set X = AsetzSsl(Verbss)
Set VerbAset = X
End Property
Function RmvEndDig$(S)
Dim J&
For J = Len(S) To 1 Step -1
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then
        RmvEndDig = Left(S, J)
        Exit Function
    End If
Next
RmvEndDig = Left(S, J)
End Function
Function Verb$(Nm)
Dim Cml$, I, LetterCml$
For Each I In CmlAy(Nm)
    Cml = I
    LetterCml = RmvDigSfx(Cml)
    If VerbAset.Has(LetterCml) Then Verb = Cml: Exit Function
Next
End Function
Property Get NormSSoVerb$()
NormSSoVerb = NormSsl(Verbss, IsDes:=True)
End Property
Function NormSsl$(Ssl$, Optional IsDes As Boolean)
NormSsl = JnSpc(AySrtQ(AwDist(SyzSS(Ssl)), IsDes:=True))
End Function

Function VerbPatn$(SSoVerb$)
Dim O$(), Verb$, I
For Each I In AsetzAy(SyzSS(SSoVerb)).Itms
    Verb = I
    PushI O, VerbPatn_(Verb)
Next
VerbPatn = QteBkt(JnVBar(O))
End Function

Function VerbPatn_$(Verb$)
ThwIf_NotVerb Verb, CSub
VerbPatn_ = Verb & "[^a-z|0-9]*"
End Function

Sub ThwIf_NotVerb(S, Fun$)
If Not IsNm(S) Then Thw Fun, "Verb must be a name", "Str", S
If Not IsAscUCas(Asc(FstChr(S))) Then Thw Fun, "Verb must started with UCase", "Str", S
End Sub

Function QteVerb$(Nm)

End Function
