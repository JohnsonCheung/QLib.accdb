Attribute VB_Name = "MxWsLnk"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxWsLnk."

Sub CrtWsnLnk()
'Ret : Fill in Go-Wsn started from act cell, down and left and crt hyp lnk to each 'go' as following:
'      if fail to do so, inf and exit: not enough spc or already have hyp lnk
'      - fill wsn : all cells fm act cell down needs to have n-ws - 1 emp cell to fill in the wsn
'                   the minus 1 is exl the cur ws.
'      - fill 'go': all cell left to act cell down also need such emp cell to fill 'Go'
'      - hyp lnk  : each go-cell, crt hyp lnk to A1 of each of the ws.  @@
Dim At As Range: Set At = ActiveCell
Dim W$():             W = AeEle(Wny(WbzRg(At)), WszRg(At).Name) ' All other wsn ept the-ws-of-@At

'== Exit & Inf if cannot Crt ===========================================================================================
If At.Column = 1 Then Debug.Print "Column cannot be 1": Exit Sub 'Exit=>
Dim R2&:           R2 = Si(W)
Dim R As Range: Set R = RgRCRC(At, 1, 0, R2, 1)
Dim Sq():          Sq = R.Value
If Not IsSqEmp(Sq) Then Debug.Print "Som cell has value in Rg[" & R.Address & "]": Exit Sub 'Exit=>
If R.Hyperlinks.Count > 0 Then Debug.Print "Som cell has HypLnk in Rg[" & R.Address & "]": Exit Sub 'Exit=>

'== Fill in Wsn / Go / CrtHypLnk =======================================================================================
Set R = At
Dim Wsn: For Each Wsn In W
    If Not IsEmpty(R.Value) Then ThwImpossible CSub ' The IsVdt is not OK
        R.Value = Wsn              ' <== Set Wsn
          Set R = RgRC(R, 1, 0)
        R.Value = "Go"             ' <== Set Go
:                 AddHypLnk R, Wsn ' <== Add HypLnk
          Set R = CellBelow(R)
Next
End Sub

