Attribute VB_Name = "Haskell_5_sort"
'Haskell_5_sort
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'====================================================================================================
    ' Function  sortIndex           昇順ソート後のインデックス配列
    ' Function  sortIndex_pred      任意の比較関数によるソート後のインデックス配列
    ' Sub       permutate           1次元配列の並べ換え
    ' Sub       permutate_back      permutate で並べ換えられた1次元配列を元の順列に戻す
    ' Function  lower_bound         ソート済み配列からのキーの検索（std::lower_boundと同じ）
    ' Function  lower_bound_pred    ソート済み配列からのキーの検索（std::lower_boundと同じ）
    ' Function  upper_bound         ソート済み配列からのキーの検索（std::upper_boundと同じ）
    ' Function  upper_bound_pred    ソート済み配列からのキーの検索（std::upper_boundと同じ）
    ' Function  equal_range         ソート済み配列からのキーの検索（std::equal_rangeと同じ）
    ' Function  equal_range_pred    ソート済み配列からのキーの検索（std::equal_rangeと同じ）
    ' Function  partition_points    ソート済み配列から条件によって区分化されている位置の一覧を得る
    ' Function  partition_points_pred
    ' Function  less_dic            述語 辞書式less
    ' Function  less_equal_dic      述語 辞書式less_equal
    ' Function  greater_dic         述語 辞書式greater
    ' Function  greater_equal_dic   述語 辞書式greater_equal
    ' Function  equal_dic           述語 辞書式equal
    ' Function  notEqual_dic        述語 辞書式notEqual
'====================================================================================================

'昇順ソート後のインデックス配列（降順ソートはこのreverseをとる）
'key_columns は2次元配列の場合のキー列指定 Array(0,2,4)
' 対象配列を実際にソートする場合は、permutate(配列, sortIndex) とするか、
' もしくはsubV(配列, sortIndex) を取る
Function sortIndex(ByRef matrix As Variant, Optional ByRef key_columns As Variant) As Variant
    Select Case Dimension(matrix)
    Case 1
        sortIndex = stdsort(matrix, 1, 0)
    Case 2
        Dim allkeyFlag As Boolean:  allkeyFlag = False
        If IsMissing(key_columns) Then
            key_columns = a_cols(matrix)
            allkeyFlag = True
        End If
        If sizeof(key_columns) = 1 Then
            sortIndex = stdsort(selectCol(matrix, key_columns(LBound(key_columns))), 1, 0)
        ElseIf allkeyFlag Then
            sortIndex = stdsort(zipC(matrix), 2, 0)
        Else
            sortIndex = stdsort(zipC(subM(matrix, , key_columns)), 2, 0)
        End If
    End Select
End Function
    Public Function p_sortIndex(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_sortIndex = make_funPointer_with_2nd_Default(AddressOf sortIndex, firstParam, secondParam)
    End Function

'任意の比較関数 comp によるソート後のインデックス配列
Function sortIndex_pred(ByRef matrix As Variant, ByRef comp As Variant) As Variant
    Select Case Dimension(matrix)
    Case 1
        sortIndex_pred = stdsort(matrix, 0, comp)
    Case 2
        sortIndex_pred = stdsort(zipC(matrix), 0, comp)
    End Select
End Function
    Public Function p_sortIndex_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_sortIndex_pred = make_funPointer(AddressOf sortIndex_pred, firstParam, secondParam, 2)
    End Function

' 1次元配列 vec の並べ換え
' s_index は sortIndex 関数、もしくは sortIndex_pred 関数の返り値を想定
' s_index に vec の範囲外の値もしくは重複があった場合の動作は未定義
' subV/subM 関数を使うより速いはず
Sub permutate(ByRef vec As Variant, ByRef s_index As Variant)
    If Dimension(vec) <> 1 Or Dimension(s_index) <> 1 Then Exit Sub
    If rowSize(vec) = 0 Or rowSize(vec) <> rowSize(s_index) Then Exit Sub
    Dim i As Long, k As Long:    k = LBound(vec)
    Dim tmp As Variant
    If VarType(vec) = VarType(Array()) Then
        ReDim tmp(LBound(vec) To UBound(vec))
        For i = LBound(s_index) To UBound(s_index) Step 1
            swapVariant tmp(k), vec(s_index(i))
            k = k + 1
        Next i
        If swapVariant(tmp, vec) = 0 Then
            For i = LBound(vec) To UBound(vec) Step 1
                swapVariant tmp(i), vec(i)
            Next i
        End If
    ElseIf IsObject(vec(LBound(vec))) Then
        tmp = vec
        For i = LBound(s_index) To UBound(s_index) Step 1
            Set vec(k) = Nothing
            Set vec(k) = tmp(s_index(i))
            Set tmp(s_index(i)) = Nothing
            k = k + 1
        Next i
    Else
        tmp = vec
        For i = LBound(s_index) To UBound(s_index) Step 1
            vec(k) = tmp(s_index(i))
            k = k + 1
        Next i
    End If
End Sub

' permutate で並べ換えられた1次元配列を元の順列に戻す
Sub permutate_back(ByRef vec As Variant, ByRef s_index As Variant)
    Dim k As Long:          k = LBound(vec)
    Dim index2() As Long:   ReDim index2(LBound(vec) To UBound(vec))
    Dim z As Variant
    For Each z In s_index
        index2(z) = k
        k = k + 1
    Next z
    permutate vec, index2
End Sub

'ソート済み配列から指定された要素以上の値が現れる最初の位置を取得
Function lower_bound(ByRef matrix As Variant, ByRef val As Variant) As Variant
    lower_bound = lower_bound_imple(matrix, val, p_less, LBound(matrix, 1), 1 + UBound(matrix, 1))
End Function
    Public Function p_lower_bound(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_lower_bound = make_funPointer(AddressOf lower_bound, firstParam, secondParam)
    End Function

'ソート済み配列から指定された要素以上の値が現れる最初の位置を取得（比較関数使用）
Function lower_bound_pred(ByRef matrix As Variant, ByRef val As Variant, ByRef pred As Variant) As Variant
    lower_bound_pred = lower_bound_imple(matrix, _
                                         val, _
                                         pred, _
                                         LBound(matrix, 1), 1 + UBound(matrix, 1))
End Function
    ' これは通常の関数オブジェクトとは異なる（比較関数のみを引数に取る）
    ' mapF_swap(p_lower_bound_pred(comp), matrix, values) という使用方法を想定
    Function p_lower_bound_pred(ByRef comp As Variant) As Variant
        p_lower_bound_pred = make_funPointer( _
                    AddressOf lower_bound_pred_zzz, _
                    Empty, _
                    make_funPointer(AddressOf makePair, yield_0, comp, 2))
    End Function
    Private Function lower_bound_pred_zzz(ByRef matrix As Variant, ByRef val_comp As Variant) As Variant
        lower_bound_pred_zzz = lower_bound_pred(matrix, val_comp(0), val_comp(1))
    End Function

'ソート済み配列から指定された要素より大きい値が現れる最初の位置を取得
Function upper_bound(ByRef matrix As Variant, ByRef val As Variant) As Variant
    upper_bound = upper_bound_imple(matrix, val, p_less, LBound(matrix, 1), 1 + UBound(matrix, 1))
End Function
    Public Function p_upper_bound(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_upper_bound = make_funPointer(AddressOf upper_bound, firstParam, secondParam)
    End Function

'ソート済み配列から指定された要素より大きい値が現れる最初の位置を取得（比較関数使用）
Function upper_bound_pred(ByRef matrix As Variant, ByRef val As Variant, ByRef pred As Variant) As Variant
    upper_bound_pred = upper_bound_imple(matrix, _
                                         val, _
                                         pred, _
                                         LBound(matrix, 1), 1 + UBound(matrix, 1))
End Function
    ' これは通常の関数オブジェクトとは異なる（比較関数のみを引数に取る）
    ' mapF_swap(p_upper_bound_pred(comp), matrix, values) という使用方法を想定
    Function p_upper_bound_pred(ByRef comp As Variant) As Variant
        p_upper_bound_pred = make_funPointer( _
                    AddressOf upper_bound_pred_zzz, _
                    Empty, _
                    make_funPointer(AddressOf makePair, yield_0, comp, 2))
    End Function
    Private Function upper_bound_pred_zzz(ByRef matrix As Variant, ByRef val_comp As Variant) As Variant
        upper_bound_pred_zzz = upper_bound_pred(matrix, val_comp(0), val_comp(1))
    End Function

'lower_boundとupper_boundの組
Function equal_range(ByRef matrix As Variant, ByRef val As Variant) As Variant
    equal_range = VBA.Array(lower_bound(matrix, val), upper_bound(matrix, val))
End Function
    Public Function p_equal_range(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal_range = make_funPointer(AddressOf equal_range, firstParam, secondParam)
    End Function

'lower_boundとupper_boundの組（比較関数使用）
Function equal_range_pred(ByRef matrix As Variant, ByRef val As Variant, ByRef pred As Variant) As Variant
    equal_range_pred = VBA.Array(lower_bound_pred(matrix, val, pred), upper_bound_pred(matrix, val, pred))
End Function
    ' これは通常の関数オブジェクトとは異なる（比較関数のみを引数に取る）
    ' mapF_swap(p_equal_range_pred(comp), matrix, values) という使用方法を想定
    Function p_equal_range_pred(ByRef comp As Variant) As Variant
        p_equal_range_pred = make_funPointer( _
                    AddressOf equal_range_pred_zzz, _
                    Empty, _
                    make_funPointer(AddressOf makePair, yield_0, comp, 2))
    End Function
    Private Function equal_range_pred_zzz(ByRef matrix As Variant, ByRef val_comp As Variant) As Variant
        equal_range_pred_zzz = equal_range_pred(matrix, val_comp(0), val_comp(1))
    End Function

'ソート済み配列から条件によって区分化されている位置の一覧を得る
Function partition_points(ByRef vec As Variant) As Variant
    partition_points = partition_points_pred(vec, p_less)
End Function

'ソート済み配列から条件によって区分化されている位置の一覧を得る
Function partition_points_pred(ByRef vec As Variant, ByRef pred As Variant) As Variant
    Dim ret As Variant
    ret = makeM(sizeof(vec) + 1)
    Dim rPos As Long:   rPos = LBound(vec)
    Dim wPos As Long:   wPos = 0
    Do
        ret(wPos) = rPos
        If UBound(vec) < rPos Then Exit Do
        rPos = upper_bound_imple(vec, vec(rPos), pred, rPos, 1 + UBound(vec))
        wPos = wPos + 1
    Loop
    ReDim Preserve ret(0 To wPos)
    swapVariant partition_points_pred, ret
End Function

'述語 辞書式less
Function less_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    less_dic = 0&
    Dim i As Long
    For i = 0 To min_fun(sizeof(a), sizeof(b)) - 1 Step 1
        If getNth_b(a, i) < getNth_b(b, i) Then
            less_dic = 1&
            Exit For
        ElseIf getNth_b(a, i) > getNth_b(b, i) Then
            Exit For
        End If
    Next i
End Function
    Function p_less_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_less_dic = make_funPointer(AddressOf less_dic, firstParam, secondParam)
    End Function

'述語 辞書式less_equal
Function less_equal_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    less_equal_dic = IIf(0 = less_dic(b, a), 1&, 0&)
End Function
    Function p_less_equal_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_less_equal_dic = make_funPointer(AddressOf less_equal_dic, firstParam, secondParam)
    End Function

'述語 辞書式greater
Function greater_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    greater_dic = less_dic(b, a)
End Function
    Function p_greater_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_greater_dic = make_funPointer(AddressOf greater_dic, firstParam, secondParam)
    End Function

'述語 辞書式greater_equal
Function greater_equal_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    greater_equal_dic = IIf(0 = less_dic(a, b), 1&, 0&)
End Function
    Function p_greater_equal_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_greater_equal_dic = make_funPointer(AddressOf greater_equal_dic, firstParam, secondParam)
    End Function

'述語 辞書式equal
Function equal_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    equal_dic = IIf(0 = less_dic(a, b) And 0 = less_dic(b, a), 1&, 0&)
End Function
    Function p_equal_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal_dic = make_funPointer(AddressOf equal_dic, firstParam, secondParam)
    End Function

'述語 辞書式notEqual
Function notEqual_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    notEqual_dic = IIf(0 = equal_dic(b, a), 1&, 0&)
End Function
    Function p_notEqual_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_notEqual_dic = make_funPointer(AddressOf notEqual_dic, firstParam, secondParam)
    End Function

'#####################
Private Function lower_bound_imple(ByRef matrix As Variant, _
                                   ByRef val As Variant, _
                                   ByRef comp As Variant, _
                                   ByVal begin_ As Long, _
                                   ByVal end_ As Long) As Long
    Dim mid_ As Long
    If end_ - begin_ < 8 Then
        Do While unbind_invoke(comp, matrix(begin_), val) 'And begin_ < end_
            begin_ = begin_ + 1
            If end_ <= begin_ Then Exit Do
        Loop
        lower_bound_imple = begin_
    Else
        mid_ = begin_ + CLng((end_ - begin_) / 2)
        If unbind_invoke(comp, matrix(mid_), val) Then
            lower_bound_imple = lower_bound_imple(matrix, val, comp, mid_, end_)
        Else
            lower_bound_imple = lower_bound_imple(matrix, val, comp, begin_, mid_)
        End If
    End If
End Function

Private Function upper_bound_imple(ByRef matrix As Variant, _
                                   ByRef val As Variant, _
                                   ByRef comp As Variant, _
                                   ByVal begin_ As Long, _
                                   ByVal end_ As Long) As Long
    Dim mid_ As Long
    If end_ - begin_ < 8 Then
        Do While 0 = unbind_invoke(comp, val, matrix(begin_)) 'And begin_ < end_
            begin_ = begin_ + 1
            If end_ <= begin_ Then Exit Do
        Loop
        upper_bound_imple = begin_
    Else
        mid_ = begin_ + CLng((end_ - begin_) / 2)
        If unbind_invoke(comp, val, matrix(mid_)) Then
            upper_bound_imple = upper_bound_imple(matrix, val, comp, begin_, mid_)
        Else
            upper_bound_imple = upper_bound_imple(matrix, val, comp, mid_, end_)
        End If
    End If
End Function
'####################

