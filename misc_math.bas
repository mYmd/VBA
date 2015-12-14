Attribute VB_Name = "misc_math"
'misc_math
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'********************************************************************
'   数学関数 (Haskell_2_stdFunにも一部の数学的関数がある)

' Function  sin_fun(p_sin)      Sin
' Function  cos_fun(p_con)      Cos
' Function  pow_fun(p_pow)      Pow
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

' Pow
Function pow_fun(ByRef x As Variant, ByRef y As Variant) As Variant
    pow_fun = x ^ y
End Function
    Function p_pow(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_pow = make_funPointer(AddressOf pow_fun, firstParam, secondParam)
    End Function

' シンプソン法による数値積分
Function integral_simpson(ByRef fun As Variant, _
                            ByVal begin_ As Double, _
                            ByVal end_ As Double, _
                            ByVal n As Long) As Double
    Dim xs As Variant, ys As Variant
    xs = mapF(p_poly(, Array((end_ - begin_) / 2 / n, begin_)), iota(0, 2 * n))
    ys = mapF(fun, xs)
    Dim constants As Variant
    ReDim constants(0 To 2 * n)
    Call fillPattern(constants, Array(2, 4))
    constants(0) = 1
    constants(2 * n) = 1
    integral_simpson = foldl1(p_plus, zipWith(p_mult, constants, ys)) * (end_ - begin_) / n / 6
End Function
