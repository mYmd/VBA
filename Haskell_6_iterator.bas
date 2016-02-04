Attribute VB_Name = "Haskell_6_iterator"
Option Explicit
'Haskell_6_iterator
'Copyright (c) 2015 mmYYmmdd

'********************************************************************
'   イテレータ
'   イテレータの生成は配列の moveがデフォルトなので注意
'====================================================================
'   Function make_iterator          1次元配列からイテレータの生成
'   Function release_iterator       イテレータの配列部分を戻して自身は解放
'   Function iterator_pos           現在のインデックス位置を取得する
'   Function iterator_advance       指しているインデックスを進める
'   Function iterator_moveTo        インデックスを任意の位置に動かす
'   Function iterator_moveTo_b      インデックスを任意の位置に動かす（対象配列のLBound基準）
'   Function iterator_get           現在のインデックスの位置の 値 を取得する
'   Function iterator_get_move      〃 move版
'   Function iterator_set           現在のインデックスの位置の 値 を設定する
'   Function iterator_set_move      〃 move版
'   Function iterator_push          現在のインデックス位置の値を設定してインデックスを進める
'   Function iterator_push_move     〃 move版
'   Function iterator_push_ex       範囲拡張しながらiterator_push
'   Function iterator_push_ex_move  〃 move版
'   Function iterator_shrink        対象配列のインデックス範囲を書込済み位置まで切り詰める
'   Function iterator_range         対象配列のインデックス範囲を取得する
'   Function iterator_check         現在のインデックスが対象配列のインデックス範囲にあるか確認する
'********************************************************************

' 1次元配列からイテレータの生成(move=trueがデフォルト)
Function make_iterator(ByRef vec As Variant, Optional ByVal move As Boolean = True) As Variant
    If Dimension(vec) = 1 Then
        Dim ret As Variant: ret = makeM(4)
        If move Then
            swapVariant ret(0), vec
        Else
            ret(0) = vec
        End If
        ret(1) = LBound(ret(0))     ' 現在位置
        ret(2) = ret(1) - 1         ' 書き込み済み位置
        swapVariant make_iterator, ret
    End If
End Function

' イテレータの配列部分を戻して自身は解放
Function release_iterator(ByRef it As Variant) As Variant
    swapVariant release_iterator, it(0)
End Function

' 現在のインデックス位置を取得する
Function iterator_pos(ByRef it As Variant) As Long
    iterator_pos = it(1)
End Function

' 指しているインデックスを進める
Function iterator_advance(ByRef it As Variant, Optional ByRef dummy As Variant) As Variant
    it(1) = it(1) + 1
    swapVariant iterator_advance, it
End Function
    Function p_iterator_advance(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_advance = make_funPointer(AddressOf iterator_advance, firstParam, secondParam)
    End Function

' インデックスを任意の位置に動かす
Function iterator_moveTo(ByRef it As Variant, ByRef pos As Variant) As Variant
    it(1) = pos
    swapVariant iterator_moveTo, it
End Function
    Function p_iterator_moveTo(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_moveTo = make_funPointer(AddressOf iterator_moveTo, firstParam, secondParam)
    End Function

' インデックスを任意の位置に動かす（対象配列のLBound基準）
Function iterator_moveTo_b(ByRef it As Variant, ByRef pos As Variant) As Variant
    If 0 <= pos Then
        it(1) = pos + LBound(it(0))
    Else
        it(1) = UBound(it(0)) + 1 + pos
    End If
    swapVariant iterator_moveTo_b, it
End Function
    Function p_iterator_moveTo_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_moveTo_b = make_funPointer(AddressOf iterator_moveTo_b, firstParam, secondParam)
    End Function

' 現在のインデックスの位置の 値 を取得する
Function iterator_get(ByRef it As Variant) As Variant
    iterator_get = it(0)(it(1))
End Function
    Function p_iterator_get(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_get = make_funPointer(AddressOf iterator_get, firstParam, secondParam)
    End Function

' iterator_get move版
Function iterator_get_move(ByRef it As Variant) As Variant
    iterator_get_move = moveVariant(it(0)(it(1)))
End Function
    Function p_iterator_get_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_get_move = make_funPointer(AddressOf iterator_get_move, firstParam, secondParam)
    End Function

' 現在のインデックスの位置の 値 を設定する
Function iterator_set(ByRef it As Variant, ByRef x As Variant) As Variant
    Dim tmp  As Variant:    tmp = x
    iterator_set = iterator_set_move(it, tmp)
End Function
    Function p_iterator_set(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_set = make_funPointer(AddressOf iterator_set, firstParam, secondParam)
    End Function

' iterator_set move版
Function iterator_set_move(ByRef it As Variant, ByRef x As Variant) As Variant
    it(0)(it(1)) = moveVariant(x)
    If it(2) < it(1) Then it(2) = it(1)
    swapVariant iterator_set_move, it
End Function
    Function p_iterator_set_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_set_move = make_funPointer(AddressOf iterator_set_move, firstParam, secondParam)
    End Function

' 現在のインデックス位置の値を設定してインデックスを進める
Function iterator_push(ByRef it As Variant, ByRef x As Variant) As Variant
    iterator_push = iterator_advance(iterator_set(it, x))
End Function
    Function p_iterator_push(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_push = make_funPointer(AddressOf iterator_push, firstParam, secondParam)
    End Function

' iterator_push move版
Function iterator_push_move(ByRef it As Variant, ByRef x As Variant) As Variant
    iterator_push_move = iterator_advance(iterator_set_move(it, x))
End Function
    Function p_iterator_push_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_push_move = make_funPointer(AddressOf iterator_push_move, firstParam, secondParam)
    End Function

' 範囲拡張しながらiterator_push
Function iterator_push_ex(ByRef it As Variant, ByRef x As Variant) As Variant
    Dim tmp As Variant:     tmp = x
    iterator_push_ex = iterator_push_move(it, tmp)
End Function
    Function p_iterator_push_ex(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_push_ex = make_funPointer(AddressOf iterator_push_ex, firstParam, secondParam)
    End Function

' iterator_push_ex move版
Function iterator_push_ex_move(ByRef it As Variant, ByRef x As Variant) As Variant
    Dim m As Long: m = max_fun(it(1), 2 * UBound(it(0)) - LBound(it(0)) + 1)
    If UBound(it(0)) < it(1) Then
        Dim tmp As Variant
        swapVariant tmp, it(0)
        ReDim Preserve tmp(LBound(tmp) To m)
        swapVariant tmp, it(0)
    End If
    iterator_push_ex_move = iterator_push_move(it, x)
End Function
    Function p_iterator_push_ex_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iterator_push_ex_move = make_funPointer(AddressOf iterator_push_ex_move, firstParam, secondParam)
    End Function

' 対象配列のインデックス範囲を書込済み位置まで切り詰める
Function iterator_shrink(ByRef it As Variant) As Variant
    Dim tmp As Variant
    swapVariant tmp, it(0)
    If LBound(tmp) <= it(2) Then
        ReDim Preserve tmp(LBound(tmp) To it(2))
    Else
        tmp = Array()
    End If
    swapVariant tmp, it(0)
    swapVariant iterator_shrink, it
End Function

' 対象配列のインデックス範囲を取得する
Function iterator_range(ByRef it As Variant) As Variant
    Dim lb As Long, UB As Long
    If Dimension(it) = 1 And sizeof(it) = 3 Then
        iterator_range = VBA.Array(LBound(it(0)), UBound(it(0)))
    Else
        iterator_range = Empty
    End If
End Function

' 現在のインデックスが対象配列のインデックス範囲にあるか確認する
Function iterator_check(ByRef it As Variant, Optional ByRef rg As Variant) As Boolean
    Dim at As Long
    If IsMissing(rg) Then rg = iterator_range(it)
    If IsEmpty(rg) Then
        iterator_check = False
    Else
        at = iterator_pos(it)
        iterator_check = (rg(0) <= at) And (at <= rg(1))
    End If
End Function
