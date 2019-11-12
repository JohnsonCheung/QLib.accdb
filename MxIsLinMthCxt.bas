Attribute VB_Name = "MxIsLinMthCxt"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIsLinMthCxt."
Const CNs$ = "AlignMth"
Function IsLinMthCxt(L) As Boolean
'Ret : True ! if @Lin should be included as Mth-Context with one of is true
'           ! #1 IsRmk and aft (rmv ' and trim) not pfx <If Stop Insp == -- .. Brw>
'           ! #2 FstChr = :
'           ! #3 SngDimColon (&IsSngDimColon)   ! a dim and only one var and aft is [:]
'           ! #4 Is Asg stmt lin (&IsLinAsg) @@
Dim Lin$: Lin = Trim(L)
Select Case True
Case HasPfx(L, "'")             ' Is Rmk
    Lin = LTrim(RmvFstChr(Lin))
    Select Case True
    Case HasPfxss(L, "If Stop Insp == -- .. Brw")     ' Don't incl if one of %PfxAy
    Case Else: IsLinMthCxt = True   ' <== Incl
    End Select
Case IsLinDimSngVarColon(Lin), IsLinAsg(Lin), FstChr(L) = ":"
    IsLinMthCxt = True              ' <== Incl
End Select
End Function

Function IsLinDimSngVarColon(L) As Boolean
'Ret true if L is Single-Dim-Colon: one V aft Dim and Colon aft DclSfx & not [For]
Dim Lin$: Lin = L
If Not ShfDim(Lin) Then Exit Function
If ShfNm(Lin) = "" Then Exit Function
ShfBkt Lin
ShfDclSfx Lin
'If HasSubStr(L, "For Each Dr In Itr(Dy") Then Stop
If FstChr(Lin) <> ":" Then Exit Function
If T1(RmvFstChr(Lin)) = "For" Then Exit Function '[Dim Dr: For ....] is False
IsLinDimSngVarColon = True
End Function

Sub Z_IsLinDimSngVarColon()
Dim L
'GoSub T0
'GoSub T1
GoSub T3
'GoSub Z
Exit Sub
T3:
    L = "Dim Dr:       For JIsDesAy() As Boolean: IsDesAy = XIsDesAy(Ay)"
    Ept = False
    GoTo Tst
T1:
    L = "Dim IsDesAy() As Boolean: IsDesAy = XIsDesAy(Ay)"
    Ept = True
    GoTo Tst
T0:
    L = "Dim A As Access.Application: Set A = DftAcs(Acs)"
    Ept = True
    GoTo Tst
Tst:
    Act = IsLinDimSngVarColon(L)
    If Act <> Ept Then Stop
    Return
Z:
    Dim A As New Aset
    For Each L In SrczP(CPj)
        L = Trim(L)
        If T1(L) = "Dim" Then
            Dim S$: S = IIf(IsLinDimSngVarColon(L), "1", "0")
            A.PushItm S & " " & L
        End If
    Next
    A.Srt.Vc
    Return
End Sub

Function IsLinAsg(L) As Boolean
'Note: [Dr(NCol) = DicId(K)] is determined as Asg-lin
Dim A$: A = LTrim(L)
ShfPfxSpc A, "Set"
If ShfDotNm(A) = "" Then Exit Function
If FstChr(A) = "(" Then
    A = AftBkt(A)
End If
IsLinAsg = T1(A) = "="
End Function
