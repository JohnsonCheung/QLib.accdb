Attribute VB_Name = "MxSampS12s"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxSampS12s."
Function SampS12s() As S12s
Dim A1$, A2$
A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df":          GoSub X
A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub X
A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df":               GoSub X
A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df":           GoSub X
A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df":            GoSub X
Exit Function
X:
    Dim B1$: B1 = RplVBar(A1)
    Dim B2$: B2 = RplVBar(A2)
    PushS12 SampS12s, S12(B1, B2)
    Return
End Function
