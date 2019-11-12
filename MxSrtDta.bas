Attribute VB_Name = "MxSrtDta"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxSrtDta."
Dim A_Dy()
Dim A_IsDesAy() As Boolean
Dim A_UC&

Function SrtDrszAllCol(A As Drs) As Drs
Dim Ixy&(): Ixy = RxyzSrtDy(A.Dy)
  SrtDrszAllCol = F_SubDrs_BySubRxy(A, Ixy)
End Function

Sub Z_SrtDrs()
Dim Drs As Drs, Act As Drs, Ept As Drs, SrtByFF$
GoSub T0
Exit Sub
T0:
    SrtByFF = "A B"
    Drs = DrszFF("A B C", DyoSSVbl("4 5 6|1 2 3|2 3 4"))
    Ept = DrszFF("A B C", DyoSSVbl("1 2 3|2 3 4|4 5 6"))
    GoTo Tst
Tst:
    Act = SrtDrs(Drs, SrtByFF)
    If Not IsEqDrs(Act, Ept) Then Stop
    Return
End Sub
Function RmvDashzDashFF$(DashFF$)
Dim Fny$(): Fny = SyzSS(DashFF)
Dim J%: For J = 0 To UB(Fny)
    Fny(J) = RmvPfx(Fny(J), "-")
Next
RmvDashzDashFF = JnSpc(Fny)
End Function

Function SrtDrsInFront(D As Drs, InFrontFF$) As Drs
Dim D1 As Drs: D1 = SelDrsInFront(D, InFrontFF)
SrtDrsInFront = SrtDrs(D1)
End Function

Function SrtDrs(D As Drs, Optional SrtByDashFF$ = "") As Drs
'@SrtByDashFF : :DashFF ! If @SrtByDashFF is blank all col of @D is used to sort.  Fld wi dash means descending
'Ret          : ! Srted Drs @@
':DashFF: :SS ! Each Term is a Nm or Dash-Nm

If NoReczDrs(D) Then SrtDrs = D: Exit Function
If SrtByDashFF = "" Then
    SrtDrs = SrtDrszAllCol(D)
    Exit Function
End If
Dim Ay$():                Ay = Ny(SrtByDashFF)           ' Each ele may have - as pfx, which mean descending
Dim Fny$():              Fny = AmRmvPfx(Ay, "-")
Dim Cxy&():              Cxy = Ixy(D.Fny, Fny)
Dim Des() As Boolean:    Des = SrtDrs__IsDesAy(Ay)
Dim Dy():                 Dy = SrtDyzCy(D.Dy, Cxy, Des)
                      SrtDrs = Drs(D.Fny, Dy)
End Function

Function BoolAyzDft(BoolAy, U&) As Boolean()
Dim A As Boolean
If IsBoolAy(BoolAy) Then
    If UB(BoolAy) = U Then
        BoolAyzDft = BoolAy
        Exit Function
    End If
End If
ReDim BoolAyzDft(U)
End Function

Sub Z_RxyzSrtDy()
Dim Dy(), IsDesAy() As Boolean
GoSub T0
GoSub T1
Exit Sub
T0:
    Dy = DyzVbl("2 a C|1 c B|3 b A")
    Ept = LngAy(1, 0, 2)
    Erase IsDesAy
    GoTo Tst
T1:
    Dy = DyzVbl("2 a C|1 c B|3 b A")
    Ept = LngAy(2, 0, 1)
    IsDesAy = BoolAyzTDot("t..")
    GoTo Tst
Tst:
    Act = RxyzSrtDy(Dy, IsDesAy)
    C
    Return
End Sub
Function RxyzSrtDy(Dy(), Optional IsDesAy) As Long()
If Si(Dy) = 0 Then Exit Function
               A_UC = UB(Dy(0))
          A_IsDesAy = BoolAyzDft(IsDesAy, A_UC)
               A_Dy = Dy
Dim L&():         L = LngSeqzU(UB(Dy))
          RxyzSrtDy = RxyzSrtDy___Srt(L)
End Function

Function SrtDyzCy(Dy(), SrtCxy&(), Optional IsDesAy) As Variant()
SrtDyzCy = AwIxy(Dy, RxyzSrtDy(SelDy(Dy, SrtCxy), IsDesAy))
End Function

Function SrtDrs__IsDesAy(Ay$()) As Boolean()
Dim I: For Each I In Ay
    PushI SrtDrs__IsDesAy, FstChr(I) = "-"
Next
End Function

Function SrtDy(Dy(), Optional IsDesAy) As Variant()
SrtDy = AwIxy(Dy, RxyzSrtDy(Dy, IsDesAy))
End Function

Function RxyzSrtDy__LE(Ixy&(), I&) As Long()
'Ret : sub-sub-of-Ixy which is LE than I
Dim KeyB: KeyB = A_Dy(I)
Dim J: For Each J In Ixy
    If RxyzSrtDy__IsLE(J, KeyB) Then PushI RxyzSrtDy__LE, J
Next
End Function

Function RxyzSrtDy__GT(Ixy&(), I&) As Long()
'Ret : sub-sub-of-Ixy which is GT than I
Dim KeyB: KeyB = A_Dy(I)
Dim J: For Each J In Ixy
    If Not RxyzSrtDy__IsLE(J, KeyB) Then PushI RxyzSrtDy__GT, J
Next
End Function

Function RxyzSrtDy__IsLE(IxA, KeyB) As Boolean
'Ret : true if @A is LE than @B
Dim KeyA: KeyA = A_Dy(IxA)
RxyzSrtDy__IsLE = IsLEzAy(KeyA, KeyB, A_IsDesAy)
End Function
Function IsLEzAy(Ay1, Ay2, IsDesAy() As Boolean) As Boolean
Dim J&: For J = 0 To UB(Ay1)
    If IsDesAy(J) Then
        If Ay1(J) < Ay2(J) Then Exit Function
        If Ay1(J) > Ay2(J) Then IsLEzAy = True: Exit Function
    Else
        If Ay1(J) > Ay2(J) Then Exit Function
        If Ay1(J) < Ay2(J) Then IsLEzAy = True: Exit Function
    End If
Next
IsLEzAy = True
End Function

Function IsGTzAy(Ay1, Ay2) As Boolean
Dim J&: For J = 0 To UB(Ay1)
    If Ay1(J) <= Ay2(J) Then Exit Function
Next
IsGTzAy = True
End Function

Function RxyzSrtDy__Swap(Ixy2&()) As Long()
Dim KeyB: KeyB = A_Dy(Ixy2(1))
If RxyzSrtDy__IsLE(Ixy2(0), KeyB) Then
    RxyzSrtDy__Swap = Ixy2
Else
    PushI RxyzSrtDy__Swap, Ixy2(1)
    PushI RxyzSrtDy__Swap, Ixy2(0)
End If
End Function

Function RxyzSrtDy___Srt(Ixy&()) As Long()
Dim O&()
    Select Case UB(Ixy)
    Case -1
    Case 0: O = Ixy
    Case 1:
        O = RxyzSrtDy__Swap(Ixy)
    Case Else
        Dim I&(): I = Ixy
        Dim P&:   P = Pop(I)
        Dim A&(): A = RxyzSrtDy__LE(I, P)
        Dim b&(): b = RxyzSrtDy__GT(I, P)
        Dim L&(): L = RxyzSrtDy___Srt(A)
        Dim H&(): H = RxyzSrtDy___Srt(b)

        PushIAy O, L
          PushI O, P
        PushIAy O, H
    End Select
RxyzSrtDy___Srt = O
End Function

Function SrtDt(A As Dt, Optional SrtByFF$ = "") As Dt
SrtDt = DtzDrs(SrtDrs(DrszDt(A), SrtByFF), A.DtNm)
End Function

Function SrtDyzC(Dy(), C&, Optional IsDes As Boolean) As Variant()
Dim Col(): Col = ColzDy(Dy, C)
Dim Ix&(): Ix = IxyzSrtAy(Col, IsDes)
Dim IFm&, ITo&, IStp%
If IsDes Then
    IFm = 0: ITo = UB(Ix): IStp = 1
Else
    IFm = UB(Ix): ITo = 0: IStp = -1
End If
Dim J&: For J = IFm To ITo Step IStp
   Push SrtDyzC, Dy(Ix(J))
Next
End Function
