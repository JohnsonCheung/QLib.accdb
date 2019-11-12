Attribute VB_Name = "MxDoMth"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CNs$ = "Do.Mth"
Const CMod$ = CLib & "MxDoMth."
Public Const FFoMth$ = "Pjn MdTy Mdn L Mdy Ty Mthn MthLin"
Public Const FFoMthe$ = "Pjn MdTy Mdn L E Mdy Ty Mthn MthLin"
Public Const FFoMthc$ = "Pjn MdTy Mdn L E Mdy Ty Mthn MthLin Mthl"
Public Const FFoPubFun$ = "Pjn MdTy Mdn L E Ty Mthn MthLin"

Function FoMthe() As String(): FoMthe = SyzSS(FFoMthe):  End Function
Function FoMthc() As String(): FoMthc = SyzSS(FFoMthc): End Function
Function FoMth() As String():  FoMth = SyzSS(FFoMth):  End Function

Function DoMthczFxa(Fxa$, Optional Xls As Excel.Application) As Drs
Dim A As Excel.Application: Set A = DftXls(Xls)
DoMthczFxa = DoMthczP(PjzFxa(Fxa))
If IsNothing(Xls) Then QuitXls Xls
End Function

Function AddColMthl(DoWith_L_E As Drs, Src$()) As Drs
Dim Dy()
    Dim IxL&, IxE&: AsgIx DoWith_L_E, "L E", IxL, IxE
    Dim Dr: For Each Dr In Itr(DoWith_L_E.Dy)
        Dim FmIx&: FmIx = Dr(IxL) - 1
        Dim ToIx&: ToIx = Dr(IxE) - 1
        Dim Mthl$: Mthl = JnCrLf(AwFT(Src, FmIx, ToIx))
        PushI Dr, Mthl
        PushI Dy, Dr
    Next
Dim Fny$(): Fny = AddEle(DoWith_L_E.Fny, "Mthl")
AddColMthl = Drs(Fny, Dy)
End Function

Function DoMthc(Src$(), DroMd()) As Drs
Dim D As Drs: D = DoMthe(Src, DroMd)
DoMthc = AddColMthl(D, Src)
End Function

Function DoMthczM(M As CodeModule) As Drs
Dim S$(): S = Src(M)
Dim N(): N = DroMdnzM(M)
DoMthczM = DoMthc(S, N)
End Function

Function DoMthcP() As Drs
DoMthcP = DoMthczP(CPj)
End Function

Sub Z_DoMthczP()
BrwDrs DoMthczP(CPj)
End Sub

Function DoMthczP(P As VBProject) As Drs
Static P_ As VBProject, O As Drs
If ObjPtr(P) <> ObjPtr(P_) Then
    Set P_ = P
    Dim C As VBComponent: For Each C In P.VBComponents
        O = AddDrs(O, DoMthczM(C.CodeModule))
    Next
End If
DoMthczP = O
End Function

Function DoMthczPjf(Pjf) As Drs
Dim V As Vbe, App, P As VBProject, PjDte As Date
OpnPjf Pjf ' Either Excel.Application or Access.Application
Set V = VbezPjf(Pjf)
Set P = PjzPjf(V, Pjf)
Select Case True
Case IsFb(Pjf):  PjDte = PjDtezAcs(CvAcs(App))
Case IsFxa(Pjf): PjDte = DtezFfn(Pjf)
Case Else: Stop
End Select
DoMthczPjf = AddCol(DoMthczP(P), "PjDte", PjDte)
If IsFb(Pjf) Then
    CvAcs(App).CloseCurrentDatabase
End If
End Function

Function DoMthczPjfy(Pjfy$()) As Drs
Dim F: For Each F In Pjfy
    ApdDrs DoMthczPjfy, DoMthczPjf(F)
Next
End Function

Function DoMthczV(V As Vbe) As Drs
Dim P As VBProject: For Each P In V.VBProjects
    Dim A As Drs: A = DoMthczP(P)
    Dim O As Drs: O = AddDrs(O, A)
Next
DoMthczV = O
End Function

Function DroMthLin(MthLin) As Variant()
With MthLinRec(MthLin)
DroMthLin = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function

Function WsoMthcP() As Worksheet
Set WsoMthcP = WszDrs(DoMthcP)
End Function

Function DoPubFunzP(P As VBProject) As Drs
DoPubFunzP = SelDrs(Dw2Eq(DoMthczP(P), "Mdy MdTy", "Pub", "Std"), FFoPubFun)
End Function

Function MthQnzMthn$(Mthn)
Dim D As Drs: D = DwEQ(DoMthP, "Mthn", Mthn)
Select Case Si(D.Dy)
Case 0: InfLin CSub, "No such Mthn[" & Mthn & "]"
Case 1:
    Dim IxMdn%: IxMdn = IxzAy(D.Fny, "Mdn")
    MthQnzMthn = D.Dy(0)(IxMdn) & "." & Mthn
Case Else
    InfLin CSub, "No then one Md has Mthn[" & Mthn & "]"
    IxMdn = IxzAy(D.Fny, "Mdn")
    Dim Dr: For Each Dr In D.Dy
        Debug.Print Dr(IxMdn) & "." & Mthn
    Next
End Select
End Function

Function DoMthP() As Drs
DoMthP = DoMthzP(CPj)
End Function

Function DoMthzP(P As VBProject) As Drs
Static O As Drs, X As Boolean
If Not X Then
    X = True
    Dim C As VBComponent: For Each C In P.VBComponents
        O = AddDrs(O, DoMthzM(C.CodeModule))
    Next
End If
DoMthzP = O
End Function

Function DoMthzM(M As CodeModule) As Drs
DoMthzM = DoMth(Src(M), DroMdnzM(M))
End Function

Sub Z_DoMtheM()
BrwDrs DoMtheM
End Sub
Function DoMtheM() As Drs
DoMtheM = DoMthezM(CMd)
End Function

Function DoMthezM(M As CodeModule) As Drs
Dim S$(): S = Src(M)
Dim D(): D = DroMdnzM(M)
DoMthezM = DoMthe(S, D)
End Function

Function DoMthe(Src$(), DroMd()) As Drs
DoMthe = AddColE(DoMth(Src, DroMd), Src)
End Function

Function DoMthM() As Drs
DoMthM = DoMthzM(CMd)
End Function

Function DoMthcM() As Drs
DoMthcM = DoMthczM(CMd)
End Function

Function AddColEzDy(DyWith_L_MthLin() As Variant, IxL&, IxMthLin&, Src$()) As Variant()
Dim Dr: For Each Dr In Itr(DyWith_L_MthLin)
    Dim Fm&: Fm = Dr(IxL) - 1
    Dim E&: E = EndLix(Src, Fm) + 1
    Dim T&: T = E - 1
    Dr = InsEleAft(Dr, E, IxL)
    PushI AddColEzDy, Dr
Next
End Function

Function AddColE(DoWith_L_MthLin As Drs, Src$()) As Drs
Dim IxL&, IxMthLin&: AsgIx DoWith_L_MthLin, "L MthLin", IxL, IxMthLin
Dim Fny$(): Fny = InsEleAft(DoWith_L_MthLin.Fny, "E", IxL)
AddColE = Drs(Fny, AddColEzDy(DoWith_L_MthLin.Dy, IxL, IxMthLin, Src))
End Function

Function DoMth(Src$(), DroMd()) As Drs
Dim Dy()
Dim Ix: For Each Ix In MthIxItr(Src)
    PushI Dy, DroMth(Ix, ContLin(Src, Ix), DroMd)
Next
DoMth = DrszFF(FFoMth, Dy)
End Function

Function DroMth(Ix, MthLin$, DroMd()) As Variant()
Dim A As Mthn3:      A = Mthn3zL(MthLin)
Dim Ty$:            Ty = A.ShtTy
Dim Mdy$:          Mdy = A.ShtMdy
Dim Mthn$:        Mthn = A.Nm
                DroMth = AddAy(DroMd, Array(Ix + 1, Mdy, Ty, Mthn, MthLin))
End Function
