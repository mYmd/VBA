Attribute VB_Name = "Haskell_5_sort"
'Haskell_5_sort
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'====================================================================================================
    ' Function  sortIndex           昇順ソート後のインデックス配列
    ' Function  sortIndex_pred      任意の比較関数によるソート後のインデックス配列
    ' Sub       permutate           1次元配列の並べ換え
    ' Function  lower_bound         ソート済み配列からのキーの検索（std::lower_boundと同じ）
    ' Function  lower_bound_pred    ソート済み配列からのキーの検索（std::lower_boundと同じ）
    ' Function  upper_bound         ソート済み配列からのキーの検索（std::upper_boundと同じ）
    ' Function  upper_bound_pred    ソート済み配列からのキーの検索（std::upper_boundと同じ）
    ' Function  equal_range         ソート済み配列からのキーの検索（std::equal_rangeと同じ）
    ' Function  equal_range_pred    ソート済み配列からのキーの検索（std::equal_rangeと同じ）
'====================================================================================================

'昇順ソート後のインデックス配列（降順ソートはこのreverseをとる）
'key_columns は2次元配列の場合のキー列指定 Array(0,2,4)
'実際にソートする場合は、permutate(配列, sortIndex) とするかもしくはsubM(配列, sortIndex) を取る
Function sortIndex(ByRef matrix As Variant, Optional ByRef key_columns As Variant) As Variant
    Select Case Dimension(matrix)
    Case 1
        sortIndex = stdsort(matrix, 1, 0)
    Case 2
        If IsMissing(key_columns) Then key_columns = a_cols(matrix)
        If sizeof(key_columns) = 1 Then
            sortIndex = stdsort(selectCol(matrix, key_columns(LBound(key_columns))), 1, 0)
        Else
            sortIndex = stdsort(foldl1(p_zip, mapF(p_selectCol(matrix), key_columns)), 2, 0)
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
        sortIndex_pred = stdsort(foldl1(p_zip, mapF(p_selectCol(matrix), a_cols(matrix))), 0, comp)
    End Select
End Function
    Public Function p_sortIndex_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_sortIndex_pred = make_funPointer(AddressOf sortIndex_pred, firstParam, secondParam)
    End Function

'1次元配列vecの並べ換え   sindexがvecの範囲外もしくは重複があった場合の動作は未定義
Sub permutate(ByRef vec As Variant, ByRef sindex As Variant)
    Dim i As Long
    Dim ret As Variant:    ReDim ret(LBound(vec) To UBound(vec))
    For i = LBound(vec) To UBound(vec) Step 1
        swapVariant ret(i), vec(sindex(i))
    Next i
    swapVariant ret, vec
End Sub


    Private Function lower_bound_imple(ByRef matrix As Variant, _
                                       ByRef val As Variant, _
                                       ByRef comp As Variant, _
                                       ByRef begin_end As Variant) As Long
        Dim begin_ As Long, end_ As Long, mid_ As Long
        begin_ = begin_end(0)
        end_ = begin_end(1)
        If end_ - begin_ < 8 Then
            Do While unbind_invoke(comp, matrix(begin_), val) 'And begin_ < end_
                begin_ = begin_ + 1
                If end_ <= begin_ Then Exit Do
            Loop
            lower_bound_imple = begin_
        Else
            mid_ = begin_ + CLng((end_ - begin_) / 2)
            If unbind_invoke(comp, matrix(mid_), val) Then
                lower_bound_imple = lower_bound_imple(matrix, val, comp, VBA.Array(mid_, end_))
            Else
                lower_bound_imple = lower_bound_imple(matrix, val, comp, VBA.Array(begin_, mid_))
            End If
        End If
    End Function
'ソート済み配列からのキーの検索（std::lower_boundと同じ）
Function lower_bound(ByRef matrix As Variant, ByRef val As Variant) As Variant
    lower_bound = lower_bound_imple(matrix, val, p_less, VBA.Array(LBound(matrix, 1), 1 + UBound(matrix, 1)))
End Function
    Public Function p_lower_bound(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_lower_bound = make_funPointer(AddressOf lower_bound, firstParam, secondParam)
    End Function

Function lower_bound_pred(ByRef matrix As Variant, ByRef val_pred As Variant) As Variant
    lower_bound_pred = lower_bound_imple(matrix, _
                                         val_pred(LBound(val_pred)), _
                                         val_pred(1 + LBound(val_pred)), _
                                         VBA.Array(LBound(matrix, 1), 1 + UBound(matrix, 1)))
End Function
    Public Function p_lower_bound_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_lower_bound_pred = make_funPointer(AddressOf lower_bound_pred, firstParam, secondParam)
    End Function

    Private Function upper_bound_imple(ByRef matrix As Variant, _
                                       ByRef val As Variant, _
                                       ByRef comp As Variant, _
                                       ByRef begin_end As Variant) As Long
        Dim begin_ As Long, end_ As Long, mid_ As Long
        begin_ = begin_end(0)
        end_ = begin_end(1)
        If end_ - begin_ < 8 Then
            Do While 0 = unbind_invoke(comp, val, matrix(begin_)) 'And begin_ < end_
                begin_ = begin_ + 1
                If end_ <= begin_ Then Exit Do
            Loop
            upper_bound_imple = begin_
        Else
            mid_ = begin_ + CLng((end_ - begin_) / 2)
            If unbind_invoke(comp, val, matrix(mid_)) Then
                upper_bound_imple = upper_bound_imple(matrix, val, comp, VBA.Array(begin_, mid_))
            Else
                upper_bound_imple = upper_bound_imple(matrix, val, comp, VBA.Array(mid_, end_))
            End If
        End If
    End Function
'ソート済み配列からのキーの検索（std::upper_boundと同じ）
Function upper_bound(ByRef matrix As Variant, ByRef val As Variant) As Variant
    upper_bound = upper_bound_imple(matrix, val, p_less, VBA.Array(LBound(matrix, 1), 1 + UBound(matrix, 1)))
End Function
    Public Function p_upper_bound(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_upper_bound = make_funPointer(AddressOf upper_bound, firstParam, secondParam)
    End Function

Function upper_bound_pred(ByRef matrix As Variant, ByRef val_pred As Variant) As Variant
    upper_bound_pred = upper_bound_imple(matrix, _
                                         val_pred(LBound(val_pred)), _
                                         val_pred(1 + LBound(val_pred)), _
                                         VBA.Array(LBound(matrix, 1), 1 + UBound(matrix, 1)))
End Function
    Public Function p_upper_bound_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_upper_bound_pred = make_funPointer(AddressOf upper_bound_pred, firstParam, secondParam)
    End Function

'ソート済み配列からのキーの検索（std::equal_rangeと同じ）
Function equal_range(ByRef matrix As Variant, ByRef val As Variant) As Variant
    equal_range = VBA.Array(lower_bound(matrix, val), upper_bound(matrix, val))
End Function
    Public Function p_equal_range(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal_range = make_funPointer(AddressOf equal_range, firstParam, secondParam)
    End Function

Function equal_range_pred(ByRef matrix As Variant, ByRef val_pred As Variant) As Variant
    equal_range_pred = VBA.Array(lower_bound_pred(matrix, val_pred), upper_bound_pred(matrix, val_pred))
End Function
    Public Function p_equal_range_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal_range_pred = make_funPointer(AddressOf equal_range_pred, firstParam, secondParam)
    End Function
