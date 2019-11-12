Attribute VB_Name = "MxAyZip"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxAyZip."

Sub Unzip(Dy2(), OAy1, OAy2)
Erase OAy1, OAy2
Dim U&: U = UB(Dy2): If U = -1 Then Exit Sub
ReDim OAy1(U)
ReDim OAy2(U)
Dim Dr, J&:: For Each Dr In Dy2
    OAy1(J) = Dr(0)
    OAy2(J) = Dr(1)
Next
End Sub

Sub Unzip3(Dy3(), OAy1, OAy2, OAy3)
Erase OAy1, OAy2, OAy3
Dim U&: U = UB(Dy3): If U = -1 Then Exit Sub
ReDim OAy1(U)
ReDim OAy2(U)
ReDim OAy3(U)
Dim Dr, J&: For Each Dr In Dy3
    OAy1(J) = Dr(0)
    OAy2(J) = Dr(1)
    OAy3(J) = Dr(2)
    J = J + 1
Next
End Sub

Sub Unzip4(Dy4(), OAy1, OAy2, OAy3, OAy4)
Erase OAy1, OAy2, OAy3, OAy4
Dim U&: U = UB(Dy4): If U = -1 Then Exit Sub
ReDim OAy1(U)
ReDim OAy2(U)
ReDim OAy3(U)
ReDim OAy4(U)
Dim Dr, J&: For Each Dr In Dy4
    OAy1(J) = Dr(0)
    OAy2(J) = Dr(1)
    OAy3(J) = Dr(2)
    OAy4(J) = Dr(3)
Next
End Sub

Function Zip(Ay1, Ay2) As Variant()
Dim U1&: U1 = UB(Ay1)
Dim U2&: U2 = UB(Ay2)
Dim U&: U = Min(U1, U2)
Dim O()
    Dim J&
    O = ResiU(O, U)
    For J = 0 To U
        O(J) = Array(Ay1(J), Ay2(J))
    Next
Zip = O
End Function

Function Zip3(Ay1, Ay2, Ay3) As Variant()
Dim U1&: U1 = UB(Ay1)
Dim U2&: U2 = UB(Ay2)
Dim U3&: U2 = UB(Ay3)
Dim U&: U = Min(U1, U2, U3)
Dim O()
    Dim J&
    O = ResiU(O, U)
    For J = 0 To U
        O(J) = Array(Ay1(J), Ay2(J), Ay3(J))
    Next
Zip3 = O
End Function

Function Zip4(Ay1, Ay2, Ay3, Ay4) As Variant()
Dim U1&: U1 = UB(Ay1)
Dim U2&: U2 = UB(Ay2)
Dim U3&: U2 = UB(Ay3)
Dim U4&: U2 = UB(Ay4)
Dim U&: U = Min(U1, U2, U3, U4)
Dim O()
    Dim J&
    O = ResiU(O, U)
    For J = 0 To U
        O(J) = Array(Ay1(J), Ay2(J), Ay3(J), Ay4(J))
    Next
Zip4 = O
End Function

Function ZipAyAp(ParamArray AyAp()) As Variant()
Dim AyAv(): If UBound(AyAp) >= 0 Then AyAv = AyAp
Dim UCol%: UCol = UB(AyAv)

Dim URow&
    Dim C%: For C = 0 To UCol
        URow = Min(URow, UB(AyAv(C)))
    Next

Dim ODy()
    Dim Dr()
    ReDim Dr(UCol)
    ReDim ODy(URow)
    Dim R&: For R = 0 To URow
        For C = 0 To UCol
            Dr(C) = Av(C)(R)
        Next
        ODy(R) = Dr
    Next
ZipAyAp = ODy
End Function

