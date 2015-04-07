Attribute VB_Name = "Haskell_1_Core"
'Haskell_1_Core
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
'   Function repeat_while       述語による条件が満たされる間繰り返し関数適用
'   Function repeat_while_not   述語による条件が満たされない間繰り返し関数適用
'   Function generate_while     述語による条件が満たされる間繰り返し関数適用の履歴を生成
'   Function generate_while_not 述語による条件が満たされない間繰り返し関数適用の履歴を生成
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
Function is_bindFun(ByRef val As Variant) As Boolean
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

' 配列 matrix の各要素でfuncによる評価結果がゼロでないものの数
Function count_if(ByRef func As Variant, ByRef matrix As Variant) As Variant
    Dim z As Variant
    count_if = 0&
    For Each z In mapF(func, matrix)
        If z <> 0 Then count_if = count_if + 1
    Next z
End Function
    Function p_count_if(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_count_if = make_funPointer(AddressOf count_if, firstParam, secondParam)
    End Function

' 述語による条件が満たされる間繰り返し関数適用
Function repeat_while(ByRef val As Variant, _
                      ByRef pred As Variant, _
                      ByRef fun As Variant, _
                      Optional ByVal N As Long = -1) As Variant
    Dim i As Long:  i = -1
    repeat_while = val
    Do While applyFun(repeat_while, pred)
        i = i + 1
        If 0 <= N And N <= i Then Exit Do
        repeat_while = applyFun(repeat_while, fun)
    Loop
End Function

' 述語による条件が満たされない間繰り返し関数適用
Function repeat_while_not(ByRef val As Variant, _
                          ByRef pred As Variant, _
                          ByRef fun As Variant, _
                          Optional ByVal N As Long = -1) As Variant
    Dim i As Long:  i = -1
    repeat_while_not = val
    Do While 0 = applyFun(repeat_while_not, pred)
        i = i + 1
        If 0 <= N And N <= i Then Exit Do
        repeat_while_not = applyFun(repeat_while_not, fun)
    Loop
End Function

' 述語による条件が満たされる間繰り返し関数適用の履歴を生成
Function generate_while(ByVal val As Variant, _
                        ByRef pred As Variant, _
                        ByRef fun As Variant, _
                        Optional ByVal N As Long = -1) As Variant
    Dim i As Long:      i = -1
    Dim ret As Variant: ReDim ret(0 To 0)
    Do While applyFun(val, pred)
        i = i + 1
        If 0 <= N And N <= i Then Exit Do
        If UBound(ret) < i Then ReDim Preserve ret(0 To i * 1)
        ret(i) = val
        val = applyFun(val, fun)
    Loop
    ReDim Preserve ret(0 To i)
    generate_while = ret
End Function

' 述語による条件が満たされない間繰り返し関数適用の履歴を生成
Function generate_while_not(ByVal val As Variant, _
                            ByRef pred As Variant, _
                            ByRef fun As Variant, _
                            Optional ByVal N As Long = -1) As Variant
    Dim i As Long:      i = -1
    Dim ret As Variant: ReDim ret(0 To 0)
    Do While 0 = applyFun(val, pred)
        i = i + 1
        If 0 <= N And N <= i Then Exit Do
        If UBound(ret) < i Then ReDim Preserve ret(0 To i * 1)
        ret(i) = val
        val = applyFun(val, fun)
    Loop
    ReDim Preserve ret(0 To i)
    generate_while_not = ret
End Function

