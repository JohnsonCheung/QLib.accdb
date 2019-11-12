Attribute VB_Name = "MxWsSize"
Option Compare Database

Property Get MaxWsCol&()
Static C&, Y As Boolean
If Not Y Then
    Y = True
    C = IIf(Xls.Version = "16.0", 16384, 255)
End If
MaxWsCol = C
End Property

Property Get MaxWsRow&()
Static R&, Y As Boolean
If Not Y Then
    Y = True
    R = IIf(Xls.Version = "16.0", 1048576, 65535)
End If
MaxWsRow = R
End Property


