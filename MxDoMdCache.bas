Attribute VB_Name = "MxDoMdCache"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMdCache."
Public Const FFoMdCache$ = "Pjn MdTy Mdn IsCached"

Function DoMdCacheP() As Drs
DoMdCacheP = AddColIsCached(DoMdnP)
End Function

Function AddColIsCached(Wi_Mdn As Drs) As Drs
Dim Dy()
    Dim IxMdn%: IxMdn = IxzAy(Wi_Mdn.Fny, "Mdn")
    Dim Dr: For Each Dr In Itr(Wi_Mdn.Dy)
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim M As CodeModule: Set M = Md(Mdn)
        Dim IsCached As Boolean: IsCached = IsCachedzM(M)
        Push Dy, AddItm(Dr, IsCached)
    Next
AddColIsCached = AddColzFFDy(Wi_Mdn, "IsCached", Dy)
End Function
