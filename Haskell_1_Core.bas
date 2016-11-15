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
'   Function moveVariant        sourceのVARIANT変数をtargetのVARIANTへmoveする
'   Function ph_0               プレースホルダ
'   Function ph_1               プレースホルダ
'   Function ph_2               プレースホルダ
'   Function yield_0            関数評価時に ph_0 を生成する
'   Function yield_1            関数評価時に ph_1 を生成する
'   Function yield_2            関数評価時に ph_2 を生成する
'   Function make_funPointer    ユーザ関数をbindファンクタ化する（関数の部分適用）
'   Function make_funPointer_with_2nd_Default  2番目の引数にデフォルト値を設定する場合
'   Function is_bindFun         bindされた関数であることの判定
'   Function bind1st            1番目の引数を再束縛する
'   Function bind2nd            2番目の引数を再束縛する
'   Sub      swap1st            第1引数のswap（大きな変数の場合に使用）
'   Sub      swap2nd            第2引数のswap（大きな変数の場合に使用）
' * Function mapF               配列の各要素に関数を適用する
'   Function mapF_swap          mapF(fun(a), m) または mapF(fun(, a), m)の構文糖
'   Function applyFun           関数適用関数
'   Function setParam           関数に引数を代入
'   Function foldl_Funs         関数合成（foldl）
'   Function scanl_Funs         関数合成（scanl）
'   Function foldr_Funs         関数合成（foldr）
'   Function scanr_Funs         関数合成（scanr）
'   Function applyFun2by2       ((x, y), (f1, f2, ...)) -> Array(f1(x, y), f2(x, y), ...)
'   Function setParam2by2       ((f1, f2, ...), (x, y)) -> Array(f1(x, y), f2(x, y), ...)
'   Function count_if           配列の各要素で述語による評価結果がゼロでないものの数
'   Function find_pred          1次元配列から条件に合致するものを検索
'   Function repeat_while       述語による条件が満たされる間繰り返し関数適用
'   Function repeat_while_not   述語による条件が満たされない間繰り返し関数適用
'   Function generate_while     述語による条件が満たされる間繰り返し関数適用の履歴を生成
'   Function generate_while_not 述語による条件が満たされない間繰り返し関数適用の履歴を生成
'   Function p_foldl            1次元配列限定のfoldl
'   Function p_foldr            1次元配列限定のfoldr
'   Function p_foldl1           1次元配列のfoldl1
'   Function p_foldr1           1次元配列のfoldr1
'   Function p_scanl            1次元配列限定のscanl
'   Function p_scanr            1次元配列限定のscanr
'   Function p_scanl1           1次元配列のscanl1
'   Function p_scanr1           1次元配列のscanr1
'   Function foldl_zipWith      zipWithをfoldlする
'   Function foldl1_zipWith     zipWithをfoldl1する
'   Function foldr_zipWith      zipWithをfoldrする
'   Function foldr1_zipWith     zipWithをfoldr1する
'   Function scanl_zipWith      zipWithをscanlする
'   Function scanr_zipWith      zipWithをscanrする
'   Function scanl1_zipWith     zipWithをscanl1する
'   Function scanr1_zipWith     zipWithをscanr1する
'***********************************************************************************

'sourceのVARIANT変数をtargetのVARIANTへmoveする
Function moveVariant(ByRef source As Variant) As Variant
    swapVariant moveVariant, source
End Function
'***********************************************************************************

'プレースホルダ（置かれた位置によって第1引数もしくは第2引数を受け取る）
Function ph_0() As Variant
    ph_0 = placeholder(0)
End Function

'プレースホルダ（第1引数を受け取る）
Function ph_1() As Variant
    ph_1 = placeholder(1)
End Function

'プレースホルダ（第2引数を受け取る）
Function ph_2() As Variant
    ph_2 = placeholder(2)
End Function

'関数評価時に ph_0 を生成する
Function yield_0() As Variant
    yield_0 = placeholder(800)
End Function

'関数評価時に ph_1 を生成する
Function yield_1() As Variant
    yield_1 = placeholder(801)
End Function

'関数評価時に ph_2 を生成する
Function yield_2() As Variant
    yield_2 = placeholder(802)
End Function

    ' Array() が IsMissing = True になることのWorkAround
    Function Is_Missing_(Optional ByRef x As Variant) As Boolean
        Is_Missing_ = IIf(IsMissing(x) And Not IsArray(x), True, False)
    End Function

'ユーザ関数をbindファンクタ化する（関数の部分適用）
'make_funPointer(func)                              引数の束縛なし
'make_funPointer(func, firstParam)                  1番目の引数を束縛
'make_funPointer(func, , secondParam)               2番目の引数を束縛
'make_funPointer(func, firstParam, secondParam)     両方の引数を束縛（遅延評価）
'functionParamPoint = 1 : firstParamが関数
'functionParamPoint = 2 : secondParamが関数
'functionParamPoint = 0 : firstParam、secondParamが値（デフォルト）
Function make_funPointer(ByVal func As LongPtr, _
                         ByRef firstParam As Variant, _
                         ByRef secondParam As Variant, _
                         Optional ByVal functionParamPoint As Long = 0) As Variant
    make_funPointer = VBA.Array(func, _
                    IIf(Is_Missing_(firstParam), yield_0, firstParam), _
                    IIf(Is_Missing_(secondParam), yield_0, secondParam), _
                    placeholder(functionParamPoint) _
                   )
End Function

'ユーザ関数をbindファンクタ化する（2番目の引数にデフォルト値を設定する場合）
Function make_funPointer_with_2nd_Default(ByVal func As LongPtr, _
                         ByRef firstParam As Variant, _
                         ByRef secondParam As Variant, _
                         Optional ByVal functionParamPoint As Long = 0) As Variant
    make_funPointer_with_2nd_Default = VBA.Array(func, _
                                 IIf(Is_Missing_(firstParam), yield_0, firstParam), _
                                 secondParam, _
                                 placeholder(functionParamPoint) _
                                )
End Function

'bindされた関数であることの判定
Function is_bindFun(ByRef val As Variant) As Boolean
    is_bindFun = False
    If Dimension(val) = 1 And sizeof(val) = 4 Then is_bindFun = is_placeholder(val(3))
End Function

'引数を再束縛する
    Private Function bind_imple(ByRef func As Variant, _
                                ByRef param As Variant, _
                                ByVal pA As Long, _
                                ByVal pB As Long, _
                                ByVal pC As Long, _
                                ByVal pD As Long) As Variant
        If is_bindFun(func) Then
            bind_imple = VBA.Array(func(0), _
                                   bind_imple(func(1), param, pA, pB, pC, pD), _
                                   bind_imple(func(2), param, pA, pB, pC, pD), _
                                   func(3))
        ElseIf is_placeholder(func) Then
            If func = placeholder(pA) Or func = placeholder(pB) Or func = placeholder(pC) Or func = placeholder(pD) Then
                bind_imple = param
            Else
                bind_imple = func
            End If
        Else
            bind_imple = func
        End If
    End Function

Function bind1st(ByRef func As Variant, ByRef firstParam As Variant, _
                    Optional ByVal ph_only As Boolean = False) As Variant
    If ph_only Then
        bind1st = VBA.Array(func(0), _
                            bind_imple(func(1), firstParam, 0, 1, 0, 1), _
                            bind_imple(func(2), firstParam, 1, 1, 1, 1), _
                            func(3))
    Else
        bind1st = VBA.Array(func(0), _
                            bind_imple(func(1), firstParam, 0, 1, 800, 801), _
                            bind_imple(func(2), firstParam, 1, 1, 801, 801), _
                            func(3))
    End If
End Function

Function bind2nd(ByRef func As Variant, ByRef secondParam As Variant, _
                    Optional ByVal ph_only As Boolean = False) As Variant
    If ph_only Then
        bind2nd = VBA.Array(func(0), _
                            bind_imple(func(1), secondParam, 2, 2, 2, 2), _
                            bind_imple(func(2), secondParam, 0, 0, 2, 2), _
                            func(3))
    Else
        bind2nd = VBA.Array(func(0), _
                            bind_imple(func(1), secondParam, 2, 2, 802, 802), _
                            bind_imple(func(2), secondParam, 0, 2, 800, 802), _
                            func(3))
    End If
End Function

'第1引数のswap（大きな変数の場合に使用）
Sub swap1st(ByRef func As Variant, ByRef firstParam As Variant)
    If is_bindFun(func) Then swapVariant func(1), firstParam
End Sub

'第2引数のswap（大きな変数の場合に使用）
Sub swap2nd(ByRef func As Variant, ByRef secondParam As Variant)
    If is_bindFun(func) Then swapVariant func(2), secondParam
End Sub

'*************************************************************************
' 配列の各要素に関数を適用する
Function mapF(ByRef func As Variant, ByRef matrix As Variant) As Variant
    mapF = mapF_imple(func, matrix)
End Function
    Function p_mapF(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_mapF = make_funPointer(AddressOf mapF, firstParam, secondParam, 1)
    End Function

' mapF(fun(a), m) または mapF(fun(, a), m)の構文糖
' ただしaはmoveされ、処理後戻される。大きな配列のコピーが避けられる。
' パターン1  mapF_swap(fun, a, m)      -> mapF(fun(a), m)
' パターン2  mapF_swap(fun, , b, m)    -> mapF(fun(, b), m)
' パターン3  mapF_swap(fun, a, , m)    -> mapF(fun(a), m)    （パターン1と同じ）
' パターン4  mapF_swap(fun, a, b, m)   -> mapF(fun(a, b), m) （禁止）
Function mapF_swap(ByRef fun As Variant, _
                        Optional ByRef x As Variant, _
                        Optional ByRef y As Variant, _
                        Optional ByRef z As Variant) As Variant
    If Is_Missing_(z) Then          ' パターン1
        swap1st fun, x
        mapF_swap = mapF(fun, y)
        swap1st fun, x
    Else
        If Is_Missing_(x) Then      ' パターン2
            swap2nd fun, y
            mapF_swap = mapF(fun, z)
            swap2nd fun, y
        ElseIf Is_Missing_(y) Then  ' パターン3
            swap1st fun, x
            mapF_swap = mapF(fun, z)
            swap1st fun, x
        Else                        ' パターン4
            '
        End If
    End If
End Function


'*************************************************************************
'関数適用関数  1引数に対して関数を適用する   関数はBind式
'1. applyFun(x     ,  Null          )     ->  x
'2. applyFun(x     ,  Empty         )     ->  x
'3. applyFun(x     , (f, a, b)      )     ->  f(a, b)
'4. applyFun(x     , (f, a) )             ->  f(a, x)
'5. applyFun(x     , (f, , b) )           ->  f(x, b)
Function applyFun(ByRef param As Variant, ByRef func As Variant) As Variant
    If IsNull(func) Or IsEmpty(func) Then
        applyFun = param
    Else
        applyFun = unbind_invoke(func, param, param)
    End If
End Function
    Function p_applyFun(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_applyFun = make_funPointer(AddressOf applyFun, firstParam, secondParam, 2)
    End Function

'関数に1引数を代入する関数
'1. setParam(f              , x     )  ->  f(x)
'2. setParam((f, a, placeholder), x )  ->  f(a, x)
'3. setParam((f, placeholder, b), x )  ->  f(x, b)
Function setParam(ByRef func As Variant, ByRef param As Variant) As Variant
    setParam = applyFun(param, func)
End Function
    Function p_setParam(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_setParam = make_funPointer(AddressOf setParam, firstParam, secondParam, 1)
    End Function

'関数合成（foldl）
Function foldl_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    foldl_Funs = foldl(p_applyFun, init, funcArray)
End Function
    Function p_foldl_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_foldl_Funs = make_funPointer(AddressOf foldl_Funs, firstParam, secondParam)
    End Function

'関数合成（scanl）
Function scanl_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    scanl_Funs = scanl(p_applyFun, init, funcArray)
End Function
    Function p_scanl_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_scanl_Funs = make_funPointer(AddressOf scanl_Funs, firstParam, secondParam)
    End Function

'関数合成（foldr）
Function foldr_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    foldr_Funs = foldr(p_setParam, init, funcArray)
End Function
    Function p_foldr_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_foldr_Funs = make_funPointer(AddressOf foldr_Funs, firstParam, secondParam)
    End Function

'関数合成（scanr）
Function scanr_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    scanr_Funs = scanr(p_setParam, init, funcArray)
End Function
    Function p_scanr_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_scanr_Funs = make_funPointer(AddressOf scanr_Funs, firstParam, secondParam)
    End Function

'((x, y), f)  に対して  f(x, y)     を返す
'((x, y), (f1, f2, ...))  に対して  Array(f1(x, y), f2(x, y), ...)     を返す
Function applyFun2by2(ByRef params As Variant, ByRef funcs As Variant) As Variant
    Dim ret As Variant, z As Variant, k As Long: k = 0
    If is_bindFun(funcs) Then
         applyFun2by2 = unbind_invoke(funcs, params(LBound(params)), params(1 + LBound(params)))
    Else
        ReDim ret(0 To sizeof(funcs) - 1)
        For Each z In funcs
            ret(k) = unbind_invoke(z, params(LBound(params)), params(1 + LBound(params)))
            k = k + 1
        Next z
        swapVariant applyFun2by2, ret
    End If
End Function
    Function p_applyFun2by2(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_applyFun2by2 = make_funPointer(AddressOf applyFun2by2, firstParam, secondParam)
    End Function

'(f, (x, y))  に対して  f(x, y)     を返す
'((f1, f2, ...), (x, y))  に対して  Array(f1(x, y), f2(x, y), ...)     を返す
Function setParam2by2(ByRef funcs As Variant, ByRef params As Variant) As Variant
    setParam2by2 = applyFun2by2(params, funcs)
End Function
    Function p_setParam2by2(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_setParam2by2 = make_funPointer(AddressOf setParam2by2, firstParam, secondParam)
    End Function

' 配列 matrix の各要素で述語による評価結果がゼロでないものの数
Function count_if(ByRef pred As Variant, ByRef matrix As Variant) As Variant
    Dim z As Variant
    count_if = 0&
    For Each z In mapF(pred, matrix)
        If z <> 0 Then count_if = count_if + 1
    Next z
End Function
    Function p_count_if(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_count_if = make_funPointer(AddressOf count_if, firstParam, secondParam, 1)
    End Function

'1次元配列から条件に合致するものを検索(最初にヒットしたインデックスを返す)
'1次元配列以外であれば返り値はEmpty、無かった場合は UBound + 1 を返す
Function find_pred(ByRef pred As Variant, ByRef vec As Variant) As Variant
    If Dimension(vec) = 1 Then
        find_pred = find_imple(pred, vec, UBound(vec) + 1)
    End If
End Function
    Function p_find_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_find_pred = make_funPointer(AddressOf find_pred, firstParam, secondParam, 1)
    End Function

' 述語による条件が満たされる間繰り返し関数適用
Function repeat_while(ByRef val As Variant, _
                      ByRef pred As Variant, _
                      ByRef fun As Variant, _
                      Optional ByVal n As Long = -1) As Variant
    repeat_while = repeat_imple(val, pred, fun, n, 0, 0)
End Function

' 述語による条件が満たされない間繰り返し関数適用
Function repeat_while_not(ByRef val As Variant, _
                          ByRef pred As Variant, _
                          ByRef fun As Variant, _
                          Optional ByVal n As Long = -1) As Variant
    repeat_while_not = repeat_imple(val, pred, fun, n, 0, 1)
End Function

' 述語による条件が満たされる間繰り返し関数適用の履歴を生成
Function generate_while(ByVal val As Variant, _
                        ByRef pred As Variant, _
                        ByRef fun As Variant, _
                        Optional ByVal n As Long = -1) As Variant
    generate_while = repeat_imple(val, pred, fun, n, 1, 0)
End Function

' 述語による条件が満たされない間繰り返し関数適用の履歴を生成
Function generate_while_not(ByVal val As Variant, _
                            ByRef pred As Variant, _
                            ByRef fun As Variant, _
                            Optional ByVal n As Long = -1) As Variant
    generate_while_not = repeat_imple(val, pred, fun, n, 1, 1)
End Function

' 1次元配列限定の foldl (p_foldl のみPublic)
    Private Function foldl_v(ByRef fun_init As Variant, ByRef vec As Variant) As Variant
        Dim fun As Variant
        fun = bind2nd(bind1st(fun_init(0), vec, True), vec, True)
        foldl_v = foldl(fun, fun_init(1), vec)
    End Function
Public Function p_foldl(ByRef fun As Variant, ByRef init As Variant) As Variant
    p_foldl = VBA.Array(AddressOf foldl_v, _
                        VBA.Array(fun, init), _
                        yield_0, _
                        placeholder)
End Function

' 1次元配列限定の foldr (p_foldr のみPublic)
    Private Function foldr_v(ByRef fun_init As Variant, ByRef vec As Variant) As Variant
        Dim fun As Variant
        fun = bind2nd(bind1st(fun_init(0), vec, True), vec, True)
        foldr_v = foldr(fun, fun_init(1), vec)
    End Function
Public Function p_foldr(ByRef fun As Variant, ByRef init As Variant) As Variant
    p_foldr = VBA.Array(AddressOf foldr_v, _
                        VBA.Array(fun, init), _
                        yield_0, _
                        placeholder)
End Function

' 1次元配列限定の foldl1 (p_foldl1 のみPublic)
    Private Function foldl1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        foldl1_v = foldl1(fun, vec)
    End Function
Public Function p_foldl1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_foldl1 = make_funPointer(AddressOf foldl1_v, firstParam, secondParam, 1)
End Function

' 1次元配列限定の foldr1 (p_foldr1 のみPublic)
    Private Function foldr1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        foldr1_v = foldr1(fun, vec)
    End Function
Public Function p_foldr1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_foldr1 = make_funPointer(AddressOf foldr1_v, firstParam, secondParam, 1)
End Function

' 1次元配列限定の scanl (p_scanl のみPublic)
    Private Function scanl_v(ByRef fun_init As Variant, ByRef vec As Variant) As Variant
        Dim fun As Variant
        fun = bind2nd(bind1st(fun_init(0), vec, True), vec, True)
        scanl_v = scanl(fun, fun_init(1), vec)
    End Function
Public Function p_scanl(ByRef fun As Variant, ByRef init As Variant) As Variant
    p_scanl = VBA.Array(AddressOf scanl_v, _
                        VBA.Array(fun, init), _
                        yield_0, _
                        placeholder)
End Function

' 1次元配列限定の scanr (p_scanr のみPublic)
    Private Function scanr_v(ByRef fun_init As Variant, ByRef vec As Variant) As Variant
        Dim fun As Variant
        fun = bind2nd(bind1st(fun_init(0), vec, True), vec, True)
        scanr_v = scanr(fun, fun_init(1), vec)
    End Function
Public Function p_scanr(ByRef fun As Variant, ByRef init As Variant) As Variant
    p_scanr = VBA.Array(AddressOf scanr_v, _
                        VBA.Array(fun, init), _
                        yield_0, _
                        placeholder)
End Function

' 1次元配列限定の scanl1 (p_scanl1 のみPublic)
    Private Function scanl1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        scanl1_v = scanl1(fun, vec)
    End Function
Public Function p_scanl1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_scanl1 = make_funPointer(AddressOf scanl1_v, firstParam, secondParam, 1)
End Function

' 1次元配列限定の scanr1 (p_scanr1 のみPublic)
    Private Function scanr1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        scanr1_v = scanr1(fun, vec)
    End Function
Public Function p_scanr1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_scanr1 = make_funPointer(AddressOf scanr1_v, firstParam, secondParam, 1)
End Function

' zipWithをfoldlする
Function foldl_zipWith(ByRef fun As Variant, ByRef init As Variant, ByRef vec As Variant) As Variant
    If LBound(vec) <= UBound(vec) Then
        foldl_zipWith = zipWith(fun, init, vec(LBound(vec)))
        Dim i As Long
        For i = LBound(vec) + 1 To UBound(vec) Step 1
            foldl_zipWith = zipWith(fun, foldl_zipWith, vec(i))
        Next i
    End If
End Function

' zipWithをfoldrする
Function foldr_zipWith(ByRef fun As Variant, ByRef init As Variant, ByRef vec As Variant) As Variant
    If LBound(vec) <= UBound(vec) Then
        foldr_zipWith = zipWith(fun, vec(UBound(vec)), init)
        Dim i As Long
        For i = UBound(vec) - 1 To LBound(vec) Step -1
            foldr_zipWith = zipWith(fun, vec(i), foldr_zipWith)
        Next i
    End If
End Function

' zipWithをfoldl1する
Function foldl1_zipWith(ByRef fun As Variant, ByRef vec As Variant) As Variant
    If LBound(vec) < UBound(vec) Then
        foldl1_zipWith = zipWith(fun, vec(LBound(vec)), vec(LBound(vec) + 1))
        Dim i As Long
        For i = LBound(vec) + 2 To UBound(vec) Step 1
            foldl1_zipWith = zipWith(fun, foldl1_zipWith, vec(i))
        Next i
    End If
End Function

' zipWithをfoldr1する
Function foldr1_zipWith(ByRef fun As Variant, ByRef vec As Variant) As Variant
    If LBound(vec) < UBound(vec) Then
        foldr1_zipWith = zipWith(fun, vec(UBound(vec) - 1), vec(UBound(vec)))
        Dim i As Long
        For i = UBound(vec) - 2 To LBound(vec) Step -1
            foldr1_zipWith = zipWith(fun, vec(i), foldr1_zipWith)
        Next i
    End If
End Function

' zipWithをscanlする
Function scanl_zipWith(ByRef fun As Variant, ByRef init As Variant, ByRef vec As Variant) As Variant
    Dim ret As Variant: ret = makeM(1 + sizeof(vec))
    Dim i As Long, k As Long: k = 0
    ret(k) = init
    For i = LBound(vec) To UBound(vec) Step 1
        k = k + 1
        ret(k) = zipWith(fun, ret(k - 1), vec(i))
    Next i
    scanl_zipWith = moveVariant(ret)
End Function

' zipWithをscanrする
Function scanr_zipWith(ByRef fun As Variant, ByRef init As Variant, ByRef vec As Variant) As Variant
    Dim ret As Variant: ret = makeM(1 + sizeof(vec))
    Dim i As Long, k As Long: k = UBound(ret)
    ret(k) = init
    For i = UBound(vec) To LBound(vec) Step -1
        k = k - 1
        ret(k) = zipWith(fun, vec(i), ret(k + 1))
    Next i
    scanr_zipWith = moveVariant(ret)
End Function

' zipWithをscanl1する
Function scanl1_zipWith(ByRef fun As Variant, ByRef vec As Variant) As Variant
    Dim ret As Variant: ret = makeM(sizeof(vec))
    Dim i As Long, k As Long: k = 0
    ret(k) = vec(LBound(vec))
    For i = LBound(vec) + 1 To UBound(vec) Step 1
        k = k + 1
        ret(k) = zipWith(fun, ret(k - 1), vec(i))
    Next i
    scanl1_zipWith = moveVariant(ret)
End Function

' zipWithをscanr1する
Function scanr1_zipWith(ByRef fun As Variant, ByRef vec As Variant) As Variant
    Dim ret As Variant: ret = makeM(sizeof(vec))
    Dim i As Long, k As Long: k = UBound(ret)
    ret(k) = vec(UBound(vec))
    For i = UBound(vec) - 1 To LBound(vec) Step -1
        k = k - 1
        ret(k) = zipWith(fun, vec(i), ret(k + 1))
    Next i
    scanr1_zipWith = moveVariant(ret)
End Function
