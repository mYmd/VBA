Attribute VB_Name = "misc_ratio"
'misc_ratio
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'**************************************************************
'　有理数の演算
'--------------------------------------------------------------
'　有理数のデータ構造はクラス化せず次の単純な配列とする
'  Array(分子, 分母) ： 分子 As Long, 分母 As Long（非負）
'  オーバーフローについて対策はしているがガードなし
'  ゼロ除算についてはガードなし
'--------------------------------------------------------------
'   Function make_ratio         :   有理数の生成
'   Function ratio2double       :   Doubleに変換
'   Function ratio2str          :   Stringに変換
'   Function ratio_plus         :   有理数の加算
'   Function ratio_negate       :   有理数の符号変更
'   Function ratio_minus        :   有理数の減算
'   Function ratio_mult         :   有理数の乗算
'   Function ratio_pow          :   有理数のベキ乗
'   Function ratio_divide       :   有理数の除算
'   Function ratio_equal        :   有理数の比較  (a = b)
'   Function ratio_not_equal    :   有理数の比較  (a <> b)
'   Function ratio_less         :   有理数の比較  (a < b)
'   Function ratio_less_equal   :   有理数の比較  (a <= b)
'   Function ratio_greater      :   有理数の比較  (a > b)
'   Function ratio_greater_equal:   有理数の比較  (a >= b)
'**************************************************************

    '最大公約数
    Public Function getGcd(ByVal a As Long, ByVal b As Long) As Long
        If a = 0 Then
            getGcd = 1
        ElseIf b = 0 Then
            getGcd = Abs(a)
        Else
            getGcd = getGcd(b, Abs(a) Mod Abs(b))
        End If
    End Function

'有理数の生成
Function make_ratio(ByRef num As Variant, ByRef den As Variant) As Variant
    Dim gcd As Long:    gcd = getGcd(num, den)
    make_ratio = VBA.Array(Sgn(num * den) * (Abs(num) \ gcd), Abs(den) \ gcd)
End Function
    Function p_make_ratio(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_make_ratio = make_funPointer(AddressOf make_ratio, firstParam, secondParam)
    End Function

'Doubleに変換
Function ratio2double(ByRef ratio As Variant, Optional ByRef secondParam As Variant) As Variant
    ratio2double = ratio(0) / ratio(1)
End Function
    Function p_ratio2double(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio2double = make_funPointer(AddressOf ratio2double, firstParam, secondParam)
    End Function

'Stringに変換
Function ratio2str(ByRef ratio As Variant, Optional ByRef secondParam As Variant) As Variant
    ratio2str = CStr(ratio(0)) & "/" & CStr(ratio(1))
End Function
    Function p_ratio2str(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio2str = make_funPointer(AddressOf ratio2str, firstParam, secondParam)
    End Function

'有理数の加算
Function ratio_plus(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
        Dim n1 As Long: n1 = ratio_1(0)
        Dim d1 As Long: d1 = ratio_1(1)
        Dim n2 As Long: n2 = ratio_2(0)
        Dim d2 As Long: d2 = ratio_2(1)
        Dim gcd As Long:    gcd = getGcd(d1, d2)
    ratio_plus = make_ratio(n1 * (d2 \ gcd) + n2 * (d1 \ gcd), d1 * (d2 \ gcd))
End Function
    Function p_ratio_plus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_plus = make_funPointer(AddressOf ratio_plus, firstParam, secondParam)
    End Function

'有理数の符号変更
Function ratio_negate(ByRef ratio As Variant, Optional ByRef dummy As Variant) As Variant
    ratio_negate = make_ratio(-ratio(0), ratio(1))
End Function
    Function p_ratio_negate(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_negate = make_funPointer(AddressOf ratio_negate, firstParam, secondParam)
    End Function

'有理数の減算
Function ratio_minus(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
    ratio_minus = ratio_plus(ratio_1, ratio_negate(ratio_2))
End Function
    Function p_ratio_minus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_minus = make_funPointer(AddressOf ratio_minus, firstParam, secondParam)
    End Function

'有理数の乗算
Function ratio_mult(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
        Dim n1 As Long: n1 = ratio_1(0)
        Dim d1 As Long: d1 = ratio_1(1)
        Dim n2 As Long: n2 = ratio_2(0)
        Dim d2 As Long: d2 = ratio_2(1)
        Dim gx As Long:    gx = getGcd(n1, d2)
        Dim gy As Long:    gy = getGcd(n2, d1)
    ratio_mult = make_ratio((n1 \ gx) * (n2 \ gy), (d2 \ gx) * (d1 \ gy))
End Function
    Function p_ratio_mult(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_mult = make_funPointer(AddressOf ratio_mult, firstParam, secondParam)
    End Function

'有理数のベキ乗
Function ratio_pow(ByRef ratio As Variant, ByRef n As Variant) As Variant
        Dim n1 As Long: n1 = ratio(0)
        Dim d1 As Long: d1 = ratio(1)
    If n = 0 Then
        ratio_pow = make_ratio(1, 1)
    ElseIf 0 < n Then
        ratio_pow = make_ratio(n1 ^ n, d1 ^ n)
    Else
        ratio_pow = make_ratio(d1 ^ -n, n1 ^ -n)
    End If
End Function
    Function p_ratio_pow(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_pow = make_funPointer(AddressOf ratio_pow, firstParam, secondParam)
    End Function

'有理数の除算
Function ratio_divide(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
        Dim n2 As Long: n2 = ratio_2(0)
        Dim d2 As Long: d2 = ratio_2(1)
    ratio_divide = ratio_mult(ratio_1, make_ratio(d2, n2))
End Function
    Function p_ratio_divide(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_divide = make_funPointer(AddressOf ratio_divide, firstParam, secondParam)
    End Function

'有理数の比較  (a = b)
Function ratio_equal(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
        Dim n1 As Long: n1 = ratio_1(0)
        Dim d1 As Long: d1 = ratio_1(1)
        Dim n2 As Long: n2 = ratio_2(0)
        Dim d2 As Long: d2 = ratio_2(1)
    ratio_equal = IIf(n1 = n2 And d1 = d2, 1, 0)
End Function
    Function p_ratio_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_equal = make_funPointer(AddressOf ratio_equal, firstParam, secondParam)
    End Function

'有理数の比較  (a <> b)
Function ratio_not_equal(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
    ratio_not_equal = IIf(ratio_equal(ratio_1, ratio_2) = 1, 0, 1)
End Function
    Function p_ratio_not_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_not_equal = make_funPointer(AddressOf ratio_not_equal, firstParam, secondParam)
    End Function

'有理数の比較  (a < b)
Function ratio_less(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
    ratio_less = IIf(ratio2double(ratio_1) < ratio2double(ratio_2), 1, 0)
End Function
    Function p_ratio_less(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_less = make_funPointer(AddressOf ratio_less, firstParam, secondParam)
    End Function

'有理数の比較  (a <= b)
Function ratio_less_equal(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
    ratio_less_equal = IIf(ratio_less(ratio_2, ratio_1), 0, 1)
End Function
    Function p_ratio_less_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_less_equal = make_funPointer(AddressOf ratio_less_equal, firstParam, secondParam)
    End Function

'有理数の比較  (a > b)
Function ratio_greater(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
    ratio_greater = ratio_less(ratio_2, ratio_1)
End Function
    Function p_ratio_greater(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_greater = make_funPointer(AddressOf ratio_greater, firstParam, secondParam)
    End Function

'有理数の比較  (a >= b)
Function ratio_greater_equal(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
    ratio_greater_equal = IIf(ratio_less(ratio_1, ratio_2), 0, 1)
End Function
    Function p_ratio_greater_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_greater_equal = make_funPointer(AddressOf ratio_greater_equal, firstParam, secondParam)
    End Function
