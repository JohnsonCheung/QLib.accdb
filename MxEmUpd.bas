Attribute VB_Name = "MxEmUpd"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxEmUpd."
Enum EmUpd
    EiRptOnly
    EiUpdAndRpt
    EiUpdOnly
End Enum

Enum EmHdr
    EiNoHdr
    EiWiHdr
End Enum

Function EmUpdStr$(Upd As EmUpd)
Dim O$
Select Case True
Case Upd = EiRptOnly: O = "*RptOnly"
Case Upd = EiUpdAndRpt: O = "*UpdAndRpt"
Case Upd = EiUpdOnly: O = "*UpdOnly"
Case Else: O = "EmUpdEr(" & Upd & ")"
End Select
EmUpdStr = O
End Function

Function IsRpt(Upd As EmUpd, Osy) As Boolean
Select Case True
Case IsRptU(Upd), IsSy(Osy): IsRpt = True
End Select
End Function

Function IsRptU(Upd As EmUpd) As Boolean
Select Case True
Case Upd = EiRptOnly, Upd = EiUpdAndRpt: IsRptU = True
End Select
End Function

Function IsUpd(Upd As EmUpd) As Boolean
Select Case True
Case Upd = EiUpdAndRpt, Upd = EiUpdOnly: IsUpd = True
End Select
End Function
