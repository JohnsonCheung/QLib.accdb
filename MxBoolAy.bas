Attribute VB_Name = "MxBoolAy"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxBoolAy."

Enum EmBoolOp
    EiOpEq = 1
    EiOpNe = 2
    EiOpAnd = 3
    EiOpOr = 4
End Enum
Enum EmEqNeOp
    EiOpEq = EmBoolOp.EiOpEq
    EiOpNe = EmBoolOp.EiOpNe
End Enum
Enum EmAndOrOp
    EiOpAnd = EmBoolOp.EiOpAnd
    EiOpOr = EmBoolOp.EiOpOr
End Enum

Function AndBoolAy(A() As Boolean) As Boolean
AndBoolAy = IsAllT(A)
End Function
Function BoolAyzTDot(TDot$) As Boolean()
Dim J%: For J = 1 To Len(TDot)
    Dim T As Boolean: T = Mid(TDot, J, 1) = "T"
    PushI BoolAyzTDot, T
Next
End Function
Function BoolAybT(IFm&, ITo&, TrueIxy&()) As Boolean()
'Fm TIxy : #True-Ixy#  :Ixy: is Ix-Ay and Ix is always running from 0.
'Ret     : #BoolAyb-fm-TrueIxy# ! where :Ayb: is Ay-base-Ix <> 0.  a bool ay of lbound=@IFm and ubound=@ITo.  Those ele pointed by @TIxy set to True.
Dim O() As Boolean: ReDim O(IFm To ITo)
Dim I: For Each I In TrueIxy
    O(I + IFm) = True
Next
BoolAybT = O
End Function

Function BoolOp(BoolOpStr$) As EmBoolOp
Dim O As EmBoolOp
Select Case UCase(BoolOpStr)
Case "AND": O = EmBoolOp.EiOpAnd
Case "OR": O = EmBoolOp.EiOpOr
Case "EQ": O = EmBoolOp.EiOpEq
Case "NE": O = EmBoolOp.EiOpNe
Case Else: Stop
End Select
BoolOp = O
End Function

Function IfStr$(IfTrue As Boolean, RetStr$)
If IfTrue Then IfStr = RetStr
End Function

Function IsAllF(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then Exit Function
Next
IsAllF = True
End Function

Function IsAllT(A() As Boolean) As Boolean
Dim I
For Each I In A
    If Not I Then Exit Function
Next
IsAllT = True
End Function

Function IsStrAndOr(A$) As Boolean
Select Case UCase(A)
Case "AND", "OR": IsStrAndOr = True
End Select
End Function

Function IsStrEqNe(A$) As Boolean
Select Case UCase(A)
Case "EQ", "NE": IsStrEqNe = True
End Select
End Function

Function IsSomF(A() As Boolean) As Boolean
Dim b: For Each b In A
    If Not b Then IsSomF = True: Exit Function
Next
End Function

Function IsSomT(A() As Boolean) As Boolean
Dim b: For Each b In A
    If b Then IsSomT = True: Exit Function
Next
End Function

Function IsVdtBoolOpStr(BoolOpStr$) As Boolean
IsVdtBoolOpStr = HitAy(BoolOpStr, BoolOpSy)
End Function

Property Get BoolOpSy() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = SyzSS("AND OR")
End If
BoolOpSy = Y
End Property
