Attribute VB_Name = "MxDoCnst"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoCnst."
Public Const FFoCnst$ = "Mdn Mdy Cnstn TyChr AftEq"

Function FoCnst() As String()
FoCnst = SyzSS(FFoCnst)
End Function

Function DoCnstP() As Drs
DoCnstP = DoCnstzP(CPj)
End Function

Function DoCnstzP(P As VBProject) As Drs
Dim O As Drs
Dim C As VBComponent: For Each C In P.VBComponents
    O = AddDrs(O, DoCnstzM(C.CodeModule))
Next
DoCnstzP = O
End Function

Function DoCnst(Dcl$(), Mdn$) As Drs
DoCnst = Drs(FoCnst, DyoCnst(Dcl, Mdn))
End Function

Function DyoCnst(Dcl$(), Mdn$) As Variant()
Dim L: For Each L In Itr(Dcl)
    PushSomSi DyoCnst, DroCnst(L, Mdn)
Next
End Function

Function DroCnst(Lin, Optional Mdn$) As Variant()
'Ret    : :Dro|EmpAv if @Lin is not a cnst-cont-lin
Dim L$: L = Lin
Dim Mdy$: Mdy = ShfMdy(L)               '<-- 1 Mdy
    Select Case Mdy
    Case "Public": Mdy = "Pub"
    Case "", "Private": Mdy = ""
    Case Else: Exit Function            '<===
    End Select

                    If Not ShfCnst(L) Then Exit Function
Dim Cnstn$: Cnstn = ShfNm(L)                '<-- 2 Nm
                    If Cnstn = "" Then Exit Function '<==
Dim TyChr$: TyChr = ShfTyChr(L)             '<-- 3 TyChr
                    If Not ShfPfx(L, " = ") Then Exit Function  '<==
          DroCnst = Array(Mdn, Mdy, Cnstn, TyChr, L)
End Function

Function DoCnstzM(M As CodeModule) As Drs
DoCnstzM = DoCnst(DclzM(M), Mdn(M))
End Function


