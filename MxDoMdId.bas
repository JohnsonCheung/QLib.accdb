Attribute VB_Name = "MxDoMdId"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMdId."
Public Const FFoMdId$ = "MdTy CLibv CNsv CModv Pjn Mdn IsCModvEr"

Function FoMdId() As String()
FoMdId = SyzSS(FFoMdId)
End Function

Function DoMdIdzP(P As VBProject) As Drs
Dim DoMdn As Drs: DoMdn = DoMdnzP(P)
Dim D1 As Drs: D1 = AddColCModv(DoMdn, P)
Dim D2 As Drs: D2 = AddColIsCModEr(D1)
DoMdIdzP = SelDrs(D2, FFoMdId)
End Function

Function AddColIsCModEr(Wi_CModv_Mdn As Drs) As Drs
Dim IxCModv%, IxMdn%: AsgIx Wi_CModv_Mdn, "CModv Mdn", IxCModv, IxMdn
Dim Dy()
    Dim Dr: For Each Dr In Itr(Wi_CModv_Mdn.Dy)
        Dim CModv$:                 CModv = Dr(IxCModv)
        Dim Mdn$:                     Mdn = Dr(IxMdn)
        Dim IsCModEr As Boolean: IsCModEr = CModv <> Mdn
        PushI Dr, IsCModEr
        PushI Dy, Dr
    Next
AddColIsCModEr = AddColzFFDy(Wi_CModv_Mdn, "IsCModvEr", Dy)
End Function

Function AddColCModv(DoMdn As Drs, P As VBProject) As Drs
Dim Dy()
    Dim IxMdn%: IxMdn = IxzAy(DoMdn.Fny, "Mdn")
    Dim Dr: For Each Dr In Itr(DoMdn.Dy)
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim M As CodeModule: Set M = P.VBComponents(Mdn).CodeModule
        PushI Dy, AddAy(Dr, DroCMod(M))
    Next
AddColCModv = AddColzFFDy(DoMdn, "CLibv CNsv CModv", Dy)
End Function

Function DoMdIdP() As Drs
DoMdIdP = DoMdIdzP(CPj)
End Function
