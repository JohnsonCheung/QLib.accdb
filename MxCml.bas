Attribute VB_Name = "MxCml"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxCml."
':Cml: :Nm #Camel#         ! :: [ Cmlf | Cmln | Cmll ] [Cmln].. ::
':Cmlf: :Nm #Camel<fat>#    ! FstNChar is UCase, rest is :CmlRestChr
':Cmln: :Nm #Camel<normal># ! FstChr is UCase, rest is :CmlRestChr
':Cmll: :Nm #Camel<lower-case># ! FstChr is Lcase, rest is :CmlRestChr
':CmlRestChr: :Chr #Camel-Rest-Chr# ! Lower-Case or _ or digit

Sub Z_ShfCml()
Dim L$, EptL$
Ept = "A"
L = "AABcDD"
EptL = "ABcDD"
GoSub Tst
Exit Sub
Tst:
    Act = ShfCml(L)
    If Act <> Ept Then Stop
    If EptL <> L Then Stop
    Return
End Sub

Function CmlAy(Nm) As String()
If Nm = "" Then Exit Function
Dim Cml$: Cml = FstChr(Nm): If Not IsLetter(Cml) Then Thw CSub, "FstChr of Nm is not Letter", "Nm", Nm
Dim IsPrvU As Boolean: IsPrvU = IsAscUCas(Asc(Cml)) ' #Is-PrvCml-LasChr-UCas#
Dim J%: For J = 2 To Len(Nm)
    Dim C$: C = Mid(Nm, J, 1)
    Dim A%: A = Asc(C)
    Dim IsCurU As Boolean: IsCurU = IsAscUCas(A)
    Select Case True
    Case IsCurU And IsPrvU: Cml = Cml & C
    Case IsCurU:            PushNB CmlAy, Cml: Cml = C
    Case IsAscNmChr(A)
        Cml = Cml & C
    Case Else: Thw CSub, "Som Chr of Nm is invalid", "Nm I Asc-Nm-I Chr-Nm-I", Nm, J, A, Chr(A)
    End Select
    IsPrvU = IsAscUCas(A)
Next
PushNB CmlAy, Cml
End Function

Function CmlAyzNy(Ny$()) As String()
Dim I, Nm$
For Each I In Itr(Ny)
    Nm = I
    PushI CmlAyzNy, CmlAy(Nm)
Next
End Function

Function Cmlss(Nm)
':Cmlss: :SS
Cmlss = JnSpc(CmlAy(Nm))
End Function

Function CmlssAy(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    PushI CmlssAy, Cmlss(N)
Next
End Function

Function CmlSetzNy(Ny$()) As Aset
Set CmlSetzNy = AsetzAy(CmlAyzNy(Ny))
End Function

Function DotCml$(Nm)
DotCml = JnDot(CmlAy(Nm))
End Function

Function FstCml$(S)
FstCml = ShfCml(CStr(S))
End Function

Function FstCmlAy(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    PushI FstCmlAy, FstCml(N)
Next
End Function

Function AscN%(S, N&)
AscN = Asc(Mid(S, N, 1))
End Function

Function IsAscCmlChr(A%) As Boolean
Select Case True
Case IsAscLetter(A), IsAscDig(A), IsAscLDash(A): IsAscCmlChr = True
End Select
End Function

Function IsAscFstCmlChr(A%) As Boolean
If IsAscLDash(A) Then Exit Function
IsAscFstCmlChr = IsAscCmlChr(A)
End Function

Function IsCmlBRK(Cml$) As Boolean
Select Case True
Case BRKCmlASet.Has(Cml), Cml = "z", IsCmlUL(Cml): IsCmlBRK = True
End Select
End Function

Function IsCmlUL(Cml$) As Boolean
Select Case True
Case Len(Cml) <> 2, Not IsAscUCas(FstAsc(Cml)), Not IsAscLCas(SndAsc(Cml))
Case Else: IsCmlUL = LCase(FstChr(Cml)) = SndChr(Cml)
End Select
End Function

Function RmvDigSfx$(S)
Dim J%
For J = Len(S) To 1 Step -1
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then RmvDigSfx = Left(S, J): Exit Function
Next
End Function

Function RmvLDashSfx$(S)
Dim J%
For J = Len(S) To 1 Step -1
    If Mid(S, J, 1) <> "_" Then RmvLDashSfx = Left(S, J): Exit Function
Next
End Function

Function Seg1ErNy() As String()
Erase XX
X "Act"
X "App"
X "Ass"
X "Ay"
X "Bar"
X "Brk"
X "C3"
X "C4"
X "Can"
X "Cell"
X "Cm"
X "Cmd"
X "Db"
X "Dbtt"
X "Dic"
X "Dy"
X "Ds"
X "Ent"
X "F"
X "Fb"
X "Fbq"
X "Fdr"
X "Fny"
X "Frm"
X "Fun"
X "Fx"
X "Git"
X "Has"
X "Lg"
X "Lgr"
X "Lnx"
X "Lo"
X "Md"
X "Min"
X "Msg"
X "Mth"
X "N"
X "O"
X "Pc"
X "Pj"
X "Ps1"
X "Pt"
X "Pth"
X "Re"
X "Res"
X "Rs"
X "Scl"
X "Sess"
X "Shp"
X "Spec"
X "Sql"
X "Sw"
X "T"
X "Tak"
X "Tim"
X "Tmp"
X "To"
X "Txtb"
X "V"
X "W"
X "Xls"
X "Y"
Seg1ErNy = XX
End Function

Function ShfCml$(OStr)
Dim J&, Fst As Boolean, Cml$, C$, A%, IsNmChr As Boolean, IsFstNmChr As Boolean
Fst = True
For J = 1 To Len(OStr)
    C = Mid(OStr, J, 1)
    A = Asc(C)
    IsNmChr = IsAscNmChr(A)
    IsFstNmChr = IsAscFstNmChr(A)
    Select Case True
    Case Fst
        Cml = C
        Fst = False
    Case IsAscUCas(A)
        If Cml <> "" Then GoTo R
        Cml = C
    Case IsAscDig(A)
        If Cml <> "" Then Cml = Cml & C
    Case IsAscLCas(A)
        Cml = Cml & C
    Case Else
        If Cml <> "" Then GoTo R
        Cml = ""
    End Select
Next
R:
    ShfCml = Cml
    OStr = Mid(OStr, J)
End Function

Function ShfCmlAy(S) As String()
Dim L$: L = S
Dim J&
While True
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    PushNB ShfCmlAy, ShfCml(L)
    If L = "" Then Exit Function
Wend
End Function

Sub Z_CmlAy()
Dim Ny$(): Ny = MthNyV
Dim N
For Each N In Ny
    If N <> Jn(CmlAy(CStr(N))) Then Stop
Next
End Sub

Function CmlRel(Ny$()) As Rel
Set CmlRel = Rel(CmlssAy(Ny))
End Function
