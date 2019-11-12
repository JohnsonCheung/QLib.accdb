Attribute VB_Name = "MxObjPP"
Option Compare Text
Option Explicit
Const CLib$ = "QItrObj."
Const CMod$ = CLib & "MxObjPP."
':PP: :Prpc-PP$ #Spc-Separated-Prpc# ! Each ele is a Prpc
':Prpc: :Dotn   #Prp-Pth# ! Prp-Pth-of-an-Object
Function OyAdd(Oy1, Oy2)
Dim O: O = Oy1
PushObjAy O, Oy2
OyAdd = O
End Function

Sub DoItoMth(ITo, ObjMth)
Dim Obj As Object
For Each Obj In ITo
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Sub DoOyMth(Oy, ObjMth)
Dim Obj
For Each Obj In Itr(Oy)
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Function FstzOyEq(Oy, Prpc, V)
Set FstzOyEq = FstzItrEq(Itr(Oy), Prpc, V)
End Function

Function AvzOyP(Oy, Prpc) As Variant()
AvzOyP = IntozOyP(EmpAv, Oy, Prpc)
End Function

Function IntozOyP(Into, Oy, Prpc)
Dim O: O = Into: Erase O
Dim Obj: For Each Obj In Itr(Oy)
    Push O, PvzC(Obj, Prpc)
Next
IntozOyP = O
End Function

Function IntAyzOyP(Oy, Prpc) As Integer()
IntAyzOyP = IntozOyP(EmpIntAy, Oy, Prpc)
End Function

Function SyzOyP(Oy, Prpc) As String()
Stop
SyzOyP = IntozOyP(EmpSy, Oy, Prpc)
End Function

Function OyeNothing(Oy)
OyeNothing = ResiU(Oy)
Dim Obj As Object
For Each Obj In Oy
    If Not IsNothing(Obj) Then PushObj OyeNothing, Obj
Next
End Function

Function OywNmPfx(Oy, NmPfx$)
Dim Obj, O
O = Oy: Erase O
For Each Obj In Itr(Oy)
    If HasPfx(Obj.Name, NmPfx) Then PushObj O, Obj
Next
OywNmPfx = O
End Function

Function OywNm(Oy, B As WhNm)
Dim Obj, O
O = Oy: Erase O
For Each Obj In Itr(Oy)
    If HitNm(Obj.Name, B) Then PushObj OywNm, Obj
Next
End Function

Function OywPredXPTrue(Oy, Xp$, P$)
Dim O, Obj As Object
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If Run(Xp, Obj, P) Then
        PushObj O, Obj
    End If
Next
OywPredXPTrue = O
End Function

Function FstzObj(Oy, Prpc$, V)
'Ret : Fst Obj in @Oy having @Prpc = @V
Dim Obj: For Each Obj In Itr(Oy)
    If PvzC(Obj, Prpc) = V Then Asg Obj, FstzObj: Exit Function
Next
End Function
Function OyzItr(Itr) As Variant()
Dim O
For Each O In Itr
    PushObj OyzItr, O
Next
End Function
Function OywIn(Oy, Prpc, InAy)
Dim Obj As Object, O
If Si(Oy) = 0 Or Si(InAy) Then OywIn = Oy: Exit Function
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If HasEle(InAy, PvzC(Obj, Prpc)) Then PushObj O, Obj
Next
OywIn = O
End Function

Function LyzObjPP(Obj As Object, PP$) As String()
Dim Prpc: For Each Prpc In SyzSS(PP)
    PushI LyzObjPP, Prpc & " " & PvzC(Obj, Prpc)
Next
End Function

Sub Z_OyDrs()
'ShwWs DrsNewWs(OyDrs(CurrentDb.TableDefs("Z_UpdSeqFld").Fields, "Name Type OrdinalPosition"))
End Sub

Sub Z_OyP_Ay()
Dim CdPanAy() As CodePane
Stop
'CdPanAy = Oy(CPj.MdAy).PrpVy("CodePane", CdPanAy)
Stop
End Sub
Sub Z_LyzObjPP()
Dim Obj As Object, PP$
GoSub T0
Exit Sub
T0:
    Set Obj = New DAO.Field
    PP = "Name Type Size"
    GoTo Tst
Tst:
    Act = LyzObjPP(Obj, PP)
    C
    Return
End Sub

