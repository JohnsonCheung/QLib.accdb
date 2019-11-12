Attribute VB_Name = "MxItrVy"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxItrVy."
Function Vvy(Itr) As Variant()
':Vvy: :Av #Variant-Val-Ay# ! V = Av = Array-of-Variant; vy = Val-Array.  It is from Itr-Ele.Value, and put into :Av
Vvy = IntoVy(EmpAv, Itr)
End Function
Function Svy(Itr) As String()
':Svy: :Sy #String-Val-Ay# ! S = Sy = String-Array; vy = Val-Array.  It is from @Itr-Ele.Value, and put into :Sy
Svy = IntoVy(EmpSy, Itr)
End Function
Function IntoVy(OIntoAy, Itr)
Erase OIntoAy
Dim Obj: For Each Obj In Itr
    Push OIntoAy, Obj.Value
Next
IntoVy = OIntoAy
End Function
