Attribute VB_Name = "MxObj"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxObj."
Const DoczP$ = "Prpc."
Const DoczPn$ = "PrpNm."
Enum EmThw
    EiThwEr
    EiNoThw
End Enum
Function IsEqObj(A, b) As Boolean
IsEqObj = ObjPtr(A) = ObjPtr(b)
End Function

Function IsEqVar(A, b) As Boolean
IsEqVar = VarPtr(A) = VarPtr(b)
End Function

Function IntozOy(OInto, Oy)
Erase OInto
Dim O, I
For Each I In Itr(Oy)
    PushObj OInto, I
Next
End Function

Function LngAyzOyPrp(Oy, Prps$) As Long()
LngAyzOyPrp = CvLngAy(IntozOyPrp(EmpLngAy, Oy, Prps))
End Function

Function IntozOyPrp(OInto, Oy, Prps$)
Dim O: O = ResiU(OInto)
Dim Obj: For Each Obj In Itr(Oy)
    Push O, PvzC(Obj, Prps)
Next
IntozOyPrp = O
End Function

Function ObjAddAy(Obj As Object, Oy)
Dim O: O = Oy
Erase O
PushObj O, Obj
PushObjAy O, Oy
ObjAddAy = O
End Function

Function ObjNm$(A)
If IsNothing(A) Then ObjNm = "#Obj Is Nothing#": Exit Function
On Error GoTo X
ObjNm = A.Name
Exit Function
X:
ObjNm = "#" & Err.Description & "#"
End Function
