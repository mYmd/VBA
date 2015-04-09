Attribute VB_Name = "Haskell_2_stdFun"
'Haskell_2_stdFun
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'********************************************************************
'   要素アクセス
' Function firstArg           1番目の引数
' Function secondArg          2番目の引数
' Function getNth             N番目の配列要素
'*********  ***********************************************************
'1番目の引数
Function firstArg(ByRef a As Variant, ByRef b As Variant) As Variant
    firstArg = a
End Function
    Function p_firstArg(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_firstArg = make_funPointer(AddressOf firstArg, firstParam, secondParam)
    End Function

'2番目の引数
Function secondArg(ByRef a As Variant, ByRef b As Variant) As Variant
    secondArg = b
End Function
    Function p_secondArg(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_secondArg = make_funPointer(AddressOf secondArg, firstParam, secondParam)
    End Function

'N番目の配列要素
Function getNth(ByRef index As Variant, ByRef matrix As Variant) As Variant
    getNth = matrix(index)
End Function
    Function p_getNth(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getNth = make_funPointer(AddressOf getNth, firstParam, secondParam)
    End Function

'********************************************************************
'     ファンクタ等
' * Function if_else        if else 選択
'   Function replaceNull    Nullを他の値に置換する
'   Function replaceEmpty   Emptyを他の値に置換する
'   Function expN           指数関数
'   Function logN           対数関数
'   Function absD           絶対値
'   Function plus           加算
'   Function minus          減算
'   Function mult           乗算
'   Function divide         除算
'   Function poly           多項式
'   Function min            min
'   Function max            max
'   Function getCLng        CLng（整数化）
'   Function str_len        Len
'   Function str_left       Left
'   Function str_right      Right
'   Function str_mid        Mid
'   Function gcm            gcm
'   Function lcm            lcm
'   Function equal          述語 Equal
'   Function notEqual       述語 Not Equal
'   Function less           述語 less
'   Function less_equal     述語 less_equal
'   Function greater        述語 greater
'   Function greater_equal  述語 greater_equal
'********************************************************************

'選択   if_else(値, [判定値(関数), 真の時の変換値(関数), 偽の時の変換値(関数)])
Function if_else(ByRef val As Variant, ByRef trans As Variant) As Variant
    Dim lb As Long
    Dim check As Boolean
    
    lb = LBound(trans)
    If is_bindFun(trans(lb)) Then
        check = applyFun(val, trans(lb))
    Else
        check = (val = trans(lb))
    End If
    If check Then
        If is_bindFun(trans(1 + lb)) Then
            if_else = applyFun(val, trans(1 + lb))
        Else
            if_else = trans(1 + lb)
        End If
    Else
        If is_bindFun(trans(2 + lb)) Then
            if_else = applyFun(val, trans(2 + lb))
        Else
            if_else = trans(2 + lb)
        End If
    End If
    If is_placeholder(if_else) Then if_else = val
End Function
    Function p_if_else(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_if_else = make_funPointer(AddressOf if_else, firstParam, secondParam)
    End Function

'Nullを他の値に置換する
Function replaceNull(ByRef x As Variant, ByRef alt As Variant) As Variant
    If IsNull(x) Then
        replaceNull = alt
    Else
        replaceNull = x
    End If
End Function
    Function p_replaceNull(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_replaceNull = make_funPointer(AddressOf replaceNull, firstParam, secondParam)
    End Function

'Emptyを他の値に置換する
Function replaceEmpty(ByRef x As Variant, ByRef alt As Variant) As Variant
    If IsEmpty(x) Then
        replaceEmpty = alt
    Else
        replaceEmpty = x
    End If
End Function
    Function p_replaceEmpty(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_replaceEmpty = make_funPointer(AddressOf replaceEmpty, firstParam, secondParam)
    End Function


'指数関数
Function expN(ByRef a As Variant, ByRef dummy As Variant) As Variant
    expN = Exp(a)
End Function
    Function p_exp(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_exp = make_funPointer(AddressOf expN, firstParam, secondParam)
    End Function

'対数関数
Function logN(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    logN = Log(a)
End Function
    Function p_log(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_log = make_funPointer(AddressOf logN, firstParam, secondParam)
    End Function

'絶対値
Function absD(ByRef val As Variant, Optional ByRef org As Variant) As Variant
    If IsMissing(org) Then org = 0
    absD = Abs(val - org)
End Function
    Function p_abs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_abs = make_funPointer(AddressOf absD, firstParam, secondParam)
    End Function

'加算
Function plus(ByRef a As Variant, ByRef b As Variant) As Variant
    plus = a + b
End Function
    Function p_plus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_plus = make_funPointer(AddressOf plus, firstParam, secondParam)
    End Function

'減算
Function minus(ByRef a As Variant, ByRef b As Variant) As Variant
    minus = a - b
End Function
    Function p_minus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_minus = make_funPointer(AddressOf minus, firstParam, secondParam)
    End Function

'乗算
Function mult(ByRef a As Variant, ByRef b As Variant) As Variant
    mult = a * b
End Function
    Function p_mult(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_mult = make_funPointer(AddressOf mult, firstParam, secondParam)
    End Function

'除算
Function divide(ByRef a As Variant, ByRef b As Variant) As Variant
    divide = a / b
End Function
    Function p_divide(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_divide = make_funPointer(AddressOf divide, firstParam, secondParam)
    End Function
    
'剰余
Function modN(ByRef a As Variant, ByRef b As Variant) As Variant
    modN = a Mod b
End Function
    Function p_mod(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_mod = make_funPointer(AddressOf modN, firstParam, secondParam)
    End Function

'多項式　（係数は高次->低次）
Function poly(ByRef x As Variant, ByRef coef As Variant) As Variant
    poly = foldr1(p_plus, zipWith(p_mult, coef, scanr(p_mult, 1, repeat(x, sizeof(coef) - 1))))
End Function
    Function p_poly(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_poly = make_funPointer(AddressOf poly, firstParam, secondParam)
    End Function

'min
Function min(ByRef a As Variant, ByRef b As Variant) As Variant
    min = IIf(a < b, a, b)
End Function
    Function p_min(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_min = make_funPointer(AddressOf min, firstParam, secondParam)
    End Function

'max
Function max(ByRef a As Variant, ByRef b As Variant) As Variant
    max = IIf(a < b, b, a)
End Function
    Function p_max(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_max = make_funPointer(AddressOf max, firstParam, secondParam)
    End Function
    
'CLng
Function getCLng(ByRef a As Variant, ByRef dummy As Variant) As Variant
    getCLng = CLng(a)
End Function
    Function p_getCLng(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getCLng = make_funPointer(AddressOf getCLng, firstParam, secondParam)
    End Function
    
'Len
Function str_len(ByRef st As Variant, ByRef dummy As Variant) As Variant
    str_len = Len(st)
End Function
    Function p_len(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getCLng = make_funPointer(AddressOf str_len, firstParam, secondParam)
    End Function
    
'Left
Function str_left(ByRef st As Variant, ByRef length As Variant) As Variant
    str_left = Left(st, length)
End Function
    Function p_left(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_left = make_funPointer(AddressOf str_left, firstParam, secondParam)
    End Function
    
'Right
Function str_right(ByRef st As Variant, ByRef length As Variant) As Variant
    str_right = Right(st, length)
End Function
    Function p_right(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_right = make_funPointer(AddressOf str_right, firstParam, secondParam)
    End Function
    
'Mid
Function str_mid(ByRef st As Variant, ByRef begin_end As Variant) As Variant
    str_mid = Mid(st, begin_end(0), begin_end(1))
End Function
    Function p_mid(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_mid = make_funPointer(AddressOf str_mid, firstParam, secondParam)
    End Function

'gcm
Function gcm(ByRef a As Variant, ByRef b As Variant) As Variant
    If a = 0 Then
        gcm = b
    ElseIf b = 0 Then
        gcm = a
    Else
        gcm = gcm(b, a Mod b)
    End If
End Function
    Function p_gcm(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_gcm = make_funPointer(AddressOf gcm, firstParam, secondParam)
    End Function
    
'lcm
Function lcm(ByRef a As Variant, ByRef b As Variant) As Variant
    lcm = a * b / gcm(a, b)
End Function
    Function p_lcm(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_lcm = make_funPointer(AddressOf lcm, firstParam, secondParam)
    End Function
    
'述語 equal
Function equal(ByRef a As Variant, ByRef b As Variant) As Variant
    equal = IIf(a = b, 1, 0)
End Function
    Function p_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal = make_funPointer(AddressOf equal, firstParam, secondParam)
    End Function

'述語 not equal
Function notEqual(ByRef a As Variant, ByRef b As Variant) As Variant
    notEqual = IIf(a = b, 0, 1)
End Function
    Function p_notEqual(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
       p_notEqual = make_funPointer(AddressOf notEqual, firstParam, secondParam)
    End Function

'述語 less
Function less(ByRef a As Variant, ByRef b As Variant) As Variant
    less = IIf(a < b, 1&, 0&)
End Function
    Function p_less(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_less = make_funPointer(AddressOf less, firstParam, secondParam)
    End Function

'述語 less_equal
Function less_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    less_equal = IIf(a <= b, 1&, 0&)
End Function
    Function p_less_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_less_equal = make_funPointer(AddressOf less_equal, firstParam, secondParam)
    End Function

'述語 greater
Function greater(ByRef a As Variant, ByRef b As Variant) As Variant
    greater = IIf(a > b, 1&, 0&)
End Function
    Function p_greater(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_greater = make_funPointer(AddressOf greater, firstParam, secondParam)
    End Function

'述語 greater_equal
Function greater_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    greater_equal = IIf(a >= b, 1&, 0&)
End Function
    Function p_greater_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_greater_equal = make_funPointer(AddressOf greater_equal, firstParam, secondParam)
    End Function
