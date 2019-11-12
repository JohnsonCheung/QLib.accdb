Attribute VB_Name = "MxDoMdId"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMdId."
Public Const FFoMdId$ = "MdTy CLibv CNsv CModv Pjn Mdn IsCModvEr"
':CNsv: :S #Cnst-CNs-Value# ! the string bet-DblQ of CnstLin-CNs of a Md
':CModv: :S #Cnst-CMod-Value# ! the string aft-rmv-sfx-[.] of bet-DblQ of CnstLin-CMod of a Md
':CLibv: :S #Cnst-CLib-Value# ! the string aft-rmv-sfx-[.] of bet-DblQ of CnstLin-CLib of a Md
Public Const FFoCMod$ = "CLibv CNsv CModv"

Function FoCMod() As String()
FoCMod = SyzSS(FFoCMod)
End Function

Function DroCMod(M As CodeModule)
Dim D$(): D = DclzM(M)
DroCMod = Array(CLibv(D), CNsv(D), CModv(D))
End Function

Function FoMdId() As String()
FoMdId = SyzSS(FFoMdId)
End Function

Function DoMdIdzP(P As VBProject) As Drs
Dim DoMdn As Drs: DoMdn = DoMdnzP(P)
Dim D1 As Drs: D1 = AddMdIdCol_CModv(DoMdn, P)
Dim D2 As Drs: D2 = AddMdIdCol_IsCModEr(D1)
DoMdIdzP = SelDrs(D2, FFoMdId)
End Function

Function AddMdIdCol_IsCModEr(Wi_CModv_Mdn As Drs) As Drs
Dim IxCModv%, IxMdn%: AsgIx Wi_CModv_Mdn, "CModv Mdn", IxCModv, IxMdn
Dim Dy()
    Dim Dr: For Each Dr In Itr(Wi_CModv_Mdn.Dy)
        Dim CModv$:                 CModv = Dr(IxCModv)
        Dim Mdn$:                     Mdn = Dr(IxMdn)
        Dim IsCModEr As Boolean: IsCModEr = CModv <> Mdn
        PushI Dr, IsCModEr
        PushI Dy, Dr
    Next
AddMdIdCol_IsCModEr = AddColzFFDy(Wi_CModv_Mdn, "IsCModvEr", Dy)
End Function

Function AddMdIdCol_CModv(Wi_Mdn As Drs, P As VBProject) As Drs
Dim Dy()
    Dim IxMdn%: IxMdn = IxzAy(Wi_Mdn.Fny, "Mdn")
    Dim Dr: For Each Dr In Itr(Wi_Mdn.Dy)
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim M As CodeModule: Set M = P.VBComponents(Mdn).CodeModule
        PushI Dy, AddAy(Dr, DroCMod(M))
    Next
AddMdIdCol_CModv = AddColzFFDy(Wi_Mdn, "CLibv CNsv CModv", Dy)
End Function

Function DoMdIdP() As Drs
DoMdIdP = DoMdIdzP(CPj)
End Function
