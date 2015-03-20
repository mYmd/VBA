Attribute VB_Name = "mapM_module"
'mapM_module
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'***********************************************************************************
'   関数型プログラム
' APIにCallbackとして渡せる関数のシグネチャは
' Function fun(ByRef x As Variant, ByRef y As Variant) As Variant
' もしくは
' Function fun(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
'===================================================================================
'   Function bindFun            関数の部分適用
'   Function make_funPointer    ユーザ関数をbindファンクタ化する
'   Function bind1st            1番目の引数を束縛
'   Function bind2nd            2番目の引数を束縛
' * Function mapF               配列の各要素に関数を適用する
'   Function applyFun           関数適用関数
'   Function setParam           関数に引数を代入
'   Function foldl_Funs         関数合成（foldl）
'   Function scanl_Funs         関数合成（scanl）
'   Function foldr_Funs         関数合成（foldr）
'   Function scanr_Funs         関数合成（scanr）
'   Function applyFun2by2
'   Function setParam2by2
'   Function count_if           配列の各要素でfuncによる評価結果がゼロでないものの数
'***********************************************************************************

'関数の部分適用
'bindFun(func)                              引数の束縛なし
'bindFun(func, firstParam)                  1番目の引数を束縛
'bindFun(func, , secondParam)               2番目の引数を束縛
'bindFun(func, firstParam, secondParam)     両方の引数を束縛（遅延評価）
Function bindFun(ByVal func As Long, _
                 Optional ByRef firstParam As Variant, _
                 Optional ByRef secondParam As Variant) As Variant
    bindFun = VBA.Array(func, _
                        IIf(IsMissing(firstParam), placeholder, firstParam), _
                        IIf(IsMissing(secondParam), placeholder, secondParam), _
                        placeholder _
                       )
End Function

'ユーザ関数をbindファンクタ化する
Function make_funPointer(ByVal func As Long, Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    If IsMissing(firstParam) And IsMissing(secondParam) Then
        make_funPointer = func
    Else
        make_funPointer = bindFun(func, firstParam, secondParam)
    End If
End Function

'bindされた関数であることの判定
Private Function is_bindFun(ByRef val As Variant) As Boolean
    is_bindFun = False
    If Dimension(val) = 1 And sizeof(val) = 4 Then is_bindFun = is_placeholder(val(3))
End Function

'プレースホルダの位置
Private Function placeholderPosition(ByRef val As Variant) As Long
    placeholderPosition = IIf(is_placeholder(val(1)), 1, 0) + IIf(is_placeholder(val(2)), 2, 0)
End Function

'1番目の引数の束縛(bindFunの構文糖)
Function bind1st(ByRef func As Variant, Optional ByRef firstParam As Variant) As Variant
    bind1st = bindFun(func, firstParam)
End Function
    Function p_bind1st(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bind1st = make_funPointer(AddressOf bind1st, firstParam, secondParam)
    End Function
    
'2番目の引数の束縛(bindFunの構文糖)
Function bind2nd(ByRef func As Variant, Optional ByRef secondParam As Variant) As Variant
    bind2nd = bindFun(func, , secondParam)
End Function
    Function p_bind2nd(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bind2nd = make_funPointer(AddressOf bind2nd, firstParam, secondParam)
    End Function

' 配列の各要素に関数を適用する
Function mapF(ByRef func As Variant, ByRef matrix As Variant) As Variant
    If IsNull(func) Or IsEmpty(func) Then
        '
    ElseIf is_bindFun(func) Then
        Select Case placeholderPosition(func)
        Case 0  '(f, a, b)
            mapF = simple_invoke(func(0), func(1), func(2))
        Case 1  ' (f, placeholder, b)
            mapF = mapL(func(0), matrix, func(2))
        Case 2  ' (f, a, placeholder)
            mapF = mapR(func(0), func(1), matrix)
        End Select
    Else
        mapF = mapL(func, matrix)
    End If
End Function
    Function p_mapF(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_mapF = make_funPointer(AddressOf mapF, firstParam, secondParam)
    End Function

'関数適用関数  引数に対して関数を適用する   関数はBind式
'1. applyFun(x     ,  Null          )     ->  x
'2. applyFun(x     ,  Empty         )     ->  x
'3. applyFun(x     ,  f             )     ->  f(x)
'4. applyFun(x     , (f, a, b)      )     ->  f(a, b)
'5. applyFun(x     , (f, a) )             ->  f(a, x)
'6. applyFun(x     , (f, , b) )           ->  f(x, b)
'7. applyFun((x, y), (f, placeholder, placeholder))  ->  f(x, y)
         ' (f, placeholder, placeholder)は bindFun(f)で作成できる
Function applyFun(ByRef param As Variant, ByRef func As Variant) As Variant
    If IsNull(func) Or IsEmpty(func) Then
        applyFun = param
    ElseIf is_bindFun(func) Then
        Select Case placeholderPosition(func)
        Case 0  '(f, a, b)
            applyFun = simple_invoke(func(0), func(1), func(2))
        Case 1  ' (f, placeholder, b)
            applyFun = simple_invoke(func(0), param, func(2))
        Case 2  ' (f, a, placeholder)
            applyFun = simple_invoke(func(0), func(1), param)
        Case 3  ' (f, placeholder, placeholder)
            applyFun = simple_invoke(func(0), param(LBound(param)), param(1 + LBound(param)))
        End Select
    Else
        applyFun = simple_invoke(func, param)
    End If
End Function
    Function p_applyFun(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_applyFun = make_funPointer(AddressOf applyFun, firstParam, secondParam)
    End Function

'関数に引数を代入する関数
'1. setParam(f              , x     )  ->  f(x)
'2. setParam((f, a, placeholder), x )  ->  f(a, x)
'3. setParam((f, placeholder, b), x )  ->  f(x, b)
'4. setParam((f, placeholder, placeholder), (x, y))  ->  f(x, y)
      ' (f, placeholder, placeholder)は bindFun(f)で作成できる
Function setParam(ByRef func As Variant, ByRef param As Variant) As Variant
    setParam = applyFun(param, func)
End Function
    Function p_setParam(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_setParam = make_funPointer(AddressOf setParam, firstParam, secondParam)
    End Function

'関数合成（foldl）
Function foldl_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    foldl_Funs = foldl(AddressOf applyFun, init, funcArray)
End Function
    Function p_foldl_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_foldl_Funs = make_funPointer(AddressOf foldl_Funs, firstParam, secondParam)
    End Function

'関数合成（scanl）
Function scanl_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    scanl_Funs = scanl(AddressOf applyFun, init, funcArray)
End Function
    Function p_scanl_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_scanl_Funs = make_funPointer(AddressOf scanl_Funs, firstParam, secondParam)
    End Function

'関数合成（foldr）
Function foldr_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    foldr_Funs = foldr(AddressOf setParam, init, funcArray)
End Function
    Function p_foldr_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_foldr_Funs = make_funPointer(AddressOf foldr_Funs, firstParam, secondParam)
    End Function

'関数合成（scanr）
Function scanr_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    scanr_Funs = scanr(AddressOf setParam, init, funcArray)
End Function
    Function p_scanr_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_scanr_Funs = make_funPointer(AddressOf scanr_Funs, firstParam, secondParam)
    End Function

'((x, y), (f1, f2))  に対して  Array(f1(x, y), f2(x, y))     を返す
Function applyFun2by2(ByRef params As Variant, ByRef funcs As Variant) As Variant
    applyFun2by2 = VBA.Array( _
          simple_invoke(funcs(LBound(funcs)), params(LBound(params)), params(1 + LBound(params))) _
        , simple_invoke(funcs(1 + LBound(funcs)), params(LBound(params)), params(1 + LBound(params))) _
                     )
End Function
    Function p_applyFun2by2(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_applyFun2by2 = make_funPointer(AddressOf applyFun2by2, firstParam, secondParam)
    End Function

'((f1, f2), (x, y))  に対して  Array(f1(x, y), f2(x, y))     を返す
Function setParam2by2(ByRef funcs As Variant, ByRef params As Variant) As Variant
    setParam2by2 = applyFun2by2(params, funcs)
End Function
    Function p_setParam2by2(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_setParam2by2 = make_funPointer(AddressOf setParam2by2, firstParam, secondParam)
    End Function

' 配列 matrix の各要素でfuncによる評価結果がゼロでないものの数   関数はBind式
Function count_if(ByRef func As Variant, ByRef matrix As Variant) As Variant
    count_if = foldl1(p_plus, mapF(p_notEqual(, 0), mapF(p_applyFun(, func), matrix)))
End Function
    Function p_count_if(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_count_if = make_funPointer(AddressOf count_if, firstParam, secondParam)
    End Function


'********************************************************************
'   要素アクセス
'   Function firstArg           1番目の引数
'   Function secondArg          2番目の引数
'   Function getNth             N番目の配列要素
'********************************************************************
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
'   Function pow            累乗
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
