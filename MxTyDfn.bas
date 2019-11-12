Attribute VB_Name = "MxTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxTyDfn."
Public Const FFoTyDfn$ = "Mdn Nm Ty Mem Rmk"

Sub Z_DoTyDfnP()
BrwDrs DoTyDfnP
End Sub

Function IsLinTyDfn(Lin) As Boolean
Dim L$: L = Lin
Dim A$: A = ShfTyDfnNm(L): If A = "" Then Exit Function
ShfDfnTy L
ShfMemNm L
If L = "" Then IsLinTyDfn = True: Exit Function
If FstChr(L) = "!" Then IsLinTyDfn = True
End Function

Function TyDfnNyP() As String()
TyDfnNyP = TyDfnNy(SrclP)
End Function

Function TyDfnNy(Srcl) As String()
TyDfnNy = TyDfnNyzS(SplitCrLf(Srcl))
End Function

Function TyDfnNyzS(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLinTyDfn(L) Then
        PushI TyDfnNyzS, RmvFstChr(T1(L))
    End If
Next
End Function
Function TyDfnNm$(Lin)
Dim T$: T = T1(Lin)
If T = "" Then Exit Function
If Fst2Chr(T) <> "':" Then Exit Function
If LasChr(T) <> ":" Then Exit Function
TyDfnNm = RmvFstChr(T)
End Function

Function IsLinTyDfnRmk(Lin) As Boolean
If FstChr(Lin) <> "'" Then Exit Function
If FstChr(LTrim(RmvFstChr(Lin))) <> "!" Then Exit Function
IsLinTyDfnRmk = True
End Function

Function IsTyDfnNm(Nm$) As Boolean
Select Case True
Case Fst2Chr(Nm) <> "':", LasChr(Nm) <> ":"
Case Else: IsTyDfnNm = True
End Select
End Function

Function IsDfnTy(Term$) As Boolean
IsDfnTy = FstChr(Term) = ":"
End Function

Function IsMemNm(Term$) As Boolean
If Len(Term) > 3 Then
    If FstChr(Term) = "#" Then
        If LasChr(Term) = "#" Then
            IsMemNm = True
        End If
    End If
End If
End Function
