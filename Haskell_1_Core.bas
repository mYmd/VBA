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
'   Function make_funPointer    ユーザ関数をbindファンクタ化する（関数の部分適用）
'   Function make_funPointer_with_2nd_Default  2番目の引数にデフォルト値を設定する場合
'   Function is_bindFun         bindされた関数であることの判定
' * Function mapF               配列の各要素に関数を適用する
'   Function applyFun           関数適用関数
'   Function setParam           関数に引数を代入
'   Function foldl_Funs         関数合成（foldl）
'   Function scanl_Funs         関数合成（scanl）
'   Function foldr_Funs         関数合成（foldr）
'   Function scanr_Funs         関数合成（scanr）
'   Function applyFun2by2
'   Function setParam2by2
'   Function count_if           配列の各要素で述語による評価結果がゼロでないものの数
'   Function find_pred          1次元配列から条件に合致するものを検索
'   Function repeat_while       述語による条件が満たされる間繰り返し関数適用
'   Function repeat_while_not   述語による条件が満たされない間繰り返し関数適用
'   Function generate_while     述語による条件が満たされる間繰り返し関数適用の履歴を生成
'   Function generate_while_not 述語による条件が満たされない間繰り返し関数適用の履歴を生成
'***********************************************************************************

'ユーザ関数をbindファンクタ化する（関数の部分適用）
'make_funPointer(func)                              引数の束縛なし
'make_funPointer(func, firstParam)                  1番目の引数を束縛
'make_funPointer(func, , secondParam)               2番目の引数を束縛
'make_funPointer(func, firstParam, secondParam)     両方の引数を束縛（遅延評価）
Function make_funPointer(ByVal func As Long, _
                         Optional ByRef firstParam As Variant, _
                         Optional ByRef secondParam As Variant) As Variant
    make_funPointer = VBA.Array(func, _
                    IIf(IsMissing(firstParam), placeholder, firstParam), _
                    IIf(IsMissing(secondParam), placeholder, secondParam), _
                    placeholder _
                   )
End Function

'ユーザ関数をbindファンクタ化する（2番目の引数にデフォルト値を設定する場合）
Function make_funPointer_with_2nd_Default(ByVal func As Long, _
                         Optional ByRef firstParam As Variant, _
                         Optional ByRef secondParam As Variant) As Variant
    make_funPointer_with_2nd_Default = VBA.Array(func, _
                                 IIf(IsMissing(firstParam), placeholder, firstParam), _
                                 secondParam, _
                                 placeholder _
                                )
End Function

'bindされた関数であることの判定
Function is_bindFun(ByRef val As Variant) As Boolean
    is_bindFun = False
    If Dimension(val) = 1 And sizeof(val) = 4 Then is_bindFun = is_placeholder(val(3))
End Function

' 配列の各要素に関数を適用する
Function mapF(ByRef func As Variant, ByRef matrix As Variant) As Variant
    mapF = mapF_imple(func, matrix)
End Function
    Function p_mapF(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_mapF = make_funPointer(AddressOf mapF, firstParam, secondParam)
    End Function

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
        applyFun = bind_invoke(func, param)
    End If
End Function
    Function p_applyFun(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_applyFun = make_funPointer(AddressOf applyFun, firstParam, secondParam)
    End Function

'関数に1引数を代入する関数
'1. setParam(f              , x     )  ->  f(x)
'2. setParam((f, a, placeholder), x )  ->  f(a, x)
'3. setParam((f, placeholder, b), x )  ->  f(x, b)
Function setParam(ByRef func As Variant, ByRef param As Variant) As Variant
    setParam = applyFun(param, func)
End Function
    Function p_setParam(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_setParam = make_funPointer(AddressOf setParam, firstParam, secondParam)
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

'((x, y), (f1, f2))  に対して  Array(f1(x, y), f2(x, y))     を返す
Function applyFun2by2(ByRef params As Variant, ByRef funcs As Variant) As Variant
    applyFun2by2 = VBA.Array( _
          unbind_invoke(funcs(LBound(funcs)), params(LBound(params)), params(1 + LBound(params))) _
        , unbind_invoke(funcs(1 + LBound(funcs)), params(LBound(params)), params(1 + LBound(params))) _
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

' 配列 matrix の各要素で述語による評価結果がゼロでないものの数
Function count_if(ByRef pred As Variant, ByRef matrix As Variant) As Variant
    Dim z As Variant
    count_if = 0&
    For Each z In mapF(pred, matrix)
        If z <> 0 Then count_if = count_if + 1
    Next z
End Function
    Function p_count_if(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_count_if = make_funPointer(AddressOf count_if, firstParam, secondParam)
    End Function

'1次元配列から条件に合致するものを検索(最初にヒットしたインデックスまたはNullを返す)
Function find_pred(ByRef pred As Variant, ByRef vec As Variant) As Variant
    If Dimension(vec) = 1 Then
        find_pred = find_imple(pred, vec, UBound(vec) + 1)
        If find_pred = UBound(vec) + 1 Then find_pred = Null
    Else
        find_pred = Null
    End If
End Function
    Function p_find_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_find_pred = make_funPointer(AddressOf find_pred, firstParam, secondParam)
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
