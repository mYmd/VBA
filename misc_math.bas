Attribute VB_Name = "misc_math"
'misc_math
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'********************************************************************
'   数学関数 (Haskell_2_stdFunにも一部の数学的関数がある)

' Function  sin_fun(p_sin)      Sin
' Function  cos_fun(p_con)      Cos
' Function  integral_simpson    シンプソン法による数値積分
'********************************************************************

' Sin
Function sin_fun(ByRef x As Variant, ByRef dummy As Variant) As Variant
    sin_fun = Sin(x)
End Function
    Function p_sin(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_sin = make_funPointer(AddressOf sin_fun, firstParam, secondParam)
    End Function

' Cos
Function cos_fun(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
    cos_fun = Cos(x)
End Function
    Function p_cos(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_cos = make_funPointer(AddressOf cos_fun, firstParam, secondParam)
    End Function

' シンプソン法による数値積分
Function integral_simpson(ByRef fun As Variant, _
                            ByVal begin_ As Double, _
                            ByVal end_ As Double, _
                            ByVal n As Long) As Double
    Dim i As Long
    Dim timesN As Variant
    Dim xs As Variant, ys As Variant
    ReDim timesN(0 To 2 * n)
    Call fillPattern(timesN, Array(2, 4))
    timesN(0) = 1
    timesN(2 * n) = 1
    xs = mapF(p_plus(begin_), mapF(p_mult((end_ - begin_) / 2 / n), iota(0, 2 * n)))
    ys = mapF(fun, xs)
    integral_simpson = foldl1(p_plus, zipWith(p_mult, timesN, ys)) * (end_ - begin_) / n / 6
End Function
