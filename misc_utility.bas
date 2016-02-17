Attribute VB_Name = "misc_utility"
'misc_utility
'Copyright (c) 2016 mmYYmmdd
Option Explicit

'*********************************************************************************
'   ユーティリティ
'*********************************************************************************
'   Function  p__n                      p_getNth_b(, n)の構文糖
'   Function  p_try                     IIf(pred(a), a', b')
'   Function  p_try_not                 IIf(Not pred(a), a', b')の構文糖
'   Function  p_try_less                p_try(p_less(p__n(0), p__n(1)), p__n(0), Null) の構文糖
'   Function  p_typename                データ型名
'   Function  p_isNumeric               IsNumeric関数
'   Function  p_format                  Format関数
'   Function  p_InStr                   InStr関数
'   Function  p_InStrRev                InStrRev関数
'   Function  p_StrConv                 StrConv関数
'   Function  cutoff_left               文字列の左側 n 文字切落
'   Function  cutoff_right              文字列の右側 n 文字切落
'   Function  separate_string           文字列の左右分離
'   Function  p_foldl1                  1次元配列のfoldl1
'   Function  p_foldr1                  1次元配列のfoldr1
'   Function  p_scanl1                  1次元配列のscanl1
'   Function  p_scanr1                  1次元配列のscanr1
'   Function  subM_R                    subM(m, 行範囲) の構文糖
'   Function  subM_R_b                  〃（LBound基準）
'   Function  subM_C                    subM(m, , 列範囲) の構文糖
'   Function  subM_C_b                  〃（LBound基準）
'   Function  selectRow_b               LBound基準のselectRow
'   Function  selectCol_b               LBound基準のselectCol
'   Sub       fillRow_b                 LBound基準のfillRow
'   Function  fillRow_b_move            LBound基準のfillRow_move
'   Sub       fillCol_b                 LBound基準のfillCol
'   Function  fillCol_b_move            LBound基準のfillCol_move
'  -----------------------------------------------------------------------------
'   Function  adjacent_op               1次元配列vecの隣接する要素間で2項操作
'   Sub       rowWise_change            2次元配列の行ごとに関数適用
'   Function  rowWise_change_move       〃moveして返す
'   Sub       columnWise_change         2次元配列の列ごとに関数適用
'   Function  colomnWise_change_move    〃moveして返す
'   Function  equal_all                 1次元配列の全要素の等値比較
'   Function  equal_all_pred            〃　述語バージョン
'  -----------------------------------------------------------------------------
'   Function  splitStr2Funs             delimiterで区切られた文字列を関数列へマッピング
'   Function  str2SummaryFun            文字列から集計関数へ変換
'   Function  str2ConvertFun            文字列から型変換関数へ変換
'  -----------------------------------------------------------------------------
'   Function  group_by_partition_points     partition_points によるGROUP-BY
'  -----------------------------------------------------------------------------
'   Function  csv2Vector                    csvファイルの1行を配列に分割
'*********************************************************************************

' p_getNth_b(, n)の構文糖
Public Function p__n(ByVal n As Long) As Variant
    p__n = p_getNth_b(, n)
End Function

' IIf(pred(a), a', b')の構文糖
Public Function p_try(ByRef pred As Variant, _
                        Optional ByRef f1 As Variant, Optional ByRef f2 As Variant) As Variant
    If IsMissing(f1) Then
        If IsMissing(f2) Then
            p_try = p_replace_0(p_if_else(, VBA.Array(pred, p_makeSole(ph_1), 0)), ph_2)
        Else
            p_try = p_replace_0(p_if_else(, VBA.Array(pred, p_makeSole(ph_1), 0)), f2)
        End If
    Else
        If IsMissing(f2) Then
            p_try = p_replace_0(p_if_else(, VBA.Array(pred, p_makeSole(f1), 0)), ph_2)
        Else
            p_try = p_replace_0(p_if_else(, VBA.Array(pred, p_makeSole(f1), 0)), f2)
        End If
    End If
End Function

' IIf(Not pred(a), a', b')の構文糖
Public Function p_try_not(ByRef pred As Variant, _
                        Optional ByRef f1 As Variant, Optional ByRef f2 As Variant) As Variant
    If IsMissing(f1) Then
        If IsMissing(f2) Then
            p_try_not = p_replace_0(p_if_else(, VBA.Array(pred, 0, p_makeSole(ph_1))), ph_2)
        Else
            p_try_not = p_replace_0(p_if_else(, VBA.Array(pred, 0, p_makeSole(ph_1))), f2)
        End If
    Else
        If IsMissing(f2) Then
            p_try_not = p_replace_0(p_if_else(, VBA.Array(pred, 0, p_makeSole(f1))), ph_2)
        Else
            p_try_not = p_replace_0(p_if_else(, VBA.Array(pred, 0, p_makeSole(f1))), f2)
        End If
    End If
End Function
    
    Private Function replace_0(ByRef x As Variant, ByRef alt As Variant) As Variant
        If IsNumeric(x) Then
            replace_0 = alt
        Else
            replace_0 = x(0)
        End If
    End Function
    Private Function p_replace_0(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_replace_0 = make_funPointer(AddressOf replace_0, firstParam, secondParam)
    End Function

' p_try(p_less(p__n(0), p__n(1)), p__n(0), Null) の構文糖
' equal_range の値を subV_if に代入するとき等に便利
Public Function p_try_less()
    p_try_less = p_try(p_less(p__n(0), p__n(1)), p__n(0), Null)
End Function

' データ型名
    Private Function typename_(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
        typename_ = TypeName(x)
    End Function
Public Function p_typename(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_typename = make_funPointer(AddressOf typename_, firstParam, secondParam)
End Function

' IsNumeric関数
        Private Function IsNumeric_(ByRef expr As Variant, Optional ByRef dummy As Variant) As Variant
            IsNumeric_ = IIf(IsNumeric(expr) And Not IsEmpty(expr) And VarType(expr) <> vbString, 1&, 0&)
        End Function
Public Function p_isNumeric(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_isNumeric = make_funPointer(AddressOf IsNumeric_, firstParam, secondParam)
End Function

' Format関数
    Private Function format_(ByRef expr As Variant, ByRef fmt As Variant) As Variant
        format_ = Format(expr, fmt)
    End Function
Public Function p_format(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_format = make_funPointer(AddressOf format_, firstParam, secondParam)
End Function

' InStr関数
     Private Function InStr_(ByRef s As Variant, ByRef expr As Variant) As Variant
        InStr_ = InStr(s, expr)
        If IsNull(InStr_) Then InStr_ = 0
    End Function
Public Function p_InStr(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_InStr = make_funPointer(AddressOf InStr_, firstParam, secondParam)
End Function

' InStrRev関数
    Private Function InStrRev_(ByRef s As Variant, ByRef expr As Variant) As Variant
        InStrRev_ = InStrRev(s, expr)
        If IsNull(InStrRev_) Then InStrRev_ = 0
    End Function
Public Function p_InStrRev(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_InStrRev = make_funPointer(AddressOf InStrRev_, firstParam, secondParam)
End Function

' StrConv関数
     Private Function StrConv_(ByRef s As Variant, ByRef expr As Variant) As Variant
        StrConv_ = StrConv(s, expr)
     End Function
Public Function p_StrConv(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_StrConv = make_funPointer(AddressOf StrConv_, firstParam, secondParam)
End Function

' 文字列の左側 n 文字切落
Function cutoff_left(ByRef expr As Variant, ByRef n As Variant) As Variant
    cutoff_left = right(expr, max_fun(0, Len(expr) - n))
End Function
    Function p_cutoff_left(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_cutoff_left = make_funPointer(AddressOf cutoff_left, firstParam, secondParam)
    End Function

' 文字列の右側 n 文字切落
Function cutoff_right(ByRef expr As Variant, ByRef n As Variant) As Variant
    cutoff_right = left(expr, max_fun(0, Len(expr) - n))
End Function
    Function p_cutoff_right(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_cutoff_right = make_funPointer(AddressOf cutoff_right, firstParam, secondParam)
    End Function

' 文字列の左右分離
Function separate_string(ByRef expr As Variant, ByRef n As Variant) As Variant
    If 0 < n Then
        separate_string = VBA.Array(left(expr, n), cutoff_left(expr, n))
    Else
        separate_string = VBA.Array(cutoff_right(expr, -n), right(expr, -n))
    End If
End Function
    Function p_separate_string(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_separate_string = make_funPointer(AddressOf separate_string, firstParam, secondParam)
    End Function

' 1次元配列限定の foldl1 (p_foldl1 のみPublic)
    Private Function foldl1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        foldl1_v = foldl1(fun, vec)
    End Function
Public Function p_foldl1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_foldl1 = make_funPointer(AddressOf foldl1_v, firstParam, secondParam)
End Function

' 1次元配列限定の foldr1 (p_foldr1 のみPublic)
    Private Function foldr1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        foldr1_v = foldr1(fun, vec)
    End Function
Public Function p_foldr1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_foldr1 = make_funPointer(AddressOf foldr1_v, firstParam, secondParam)
End Function

' 1次元配列限定の scanl1 (p_scanl1 のみPublic)
    Private Function scanl1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        scanl1_v = scanl1(fun, vec)
    End Function
Public Function p_scanl1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_scanl1 = make_funPointer(AddressOf scanl1_v, firstParam, secondParam)
End Function

' 1次元配列限定の scanr1 (p_scanr1 のみPublic)
    Private Function scanr1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        scanr1_v = scanr1(fun, vec)
    End Function
Public Function p_scanr1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_scanr1 = make_funPointer(AddressOf scanr1_v, firstParam, secondParam)
End Function

' subM(m, 行範囲) の構文糖
Public Function subM_R(ByRef m As Variant, ByRef rRange As Variant) As Variant
    If IsArray(rRange) Then
        subM_R = subM(m, rRange)
    Else
        subM_R = subM(m, VBA.Array(rRange))
    End If
End Function
    Public Function p_subM_R(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subM_R = make_funPointer(AddressOf subM_R, firstParam, secondParam)
    End Function

' subM(m, 行範囲) の構文糖（LBound基準）
Public Function subM_R_b(ByRef m As Variant, ByRef rRange As Variant) As Variant
    Dim range_b As Variant
    range_b = mapF(p_if_else(, Array(p_less_equal(0), p_plus(LBound(m, 1)), p_plus(1 + UBound(m, 1)))), rRange)
    subM_R_b = subM_R(m, range_b)
End Function
    Public Function p_subM_R_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subM_R_b = make_funPointer(AddressOf subM_R_b, firstParam, secondParam)
    End Function

' subM(m, , 列範囲) の構文糖
Public Function subM_C(ByRef m As Variant, ByRef cRange As Variant) As Variant
    If IsArray(cRange) Then
        subM_C = subM(m, , cRange)
    Else
        subM_C = subM(m, , VBA.Array(cRange))
    End If
End Function
    Public Function p_subM_C(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subM_C = make_funPointer(AddressOf subM_C, firstParam, secondParam)
    End Function

' subM(m, , 列範囲) の構文糖（LBound基準）
Public Function subM_C_b(ByRef m As Variant, ByRef cRange As Variant) As Variant
    Dim range_b As Variant
    range_b = mapF(p_if_else(, Array(p_less_equal(0), p_plus(LBound(m, 2)), p_plus(1 + UBound(m, 2)))), cRange)
    subM_C_b = subM_C(m, range_b)
End Function
    Public Function p_subM_C_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subM_C_b = make_funPointer(AddressOf subM_C_b, firstParam, secondParam)
    End Function

'特定行の取得（LBound基準）
'index < 0 の場合は後ろから取得
Public Function selectRow_b(ByRef matrix As Variant, ByRef i As Variant) As Variant
    If 0 <= i Then
        selectRow_b = selectRow(matrix, i + LBound(matrix, 1))
    Else
        selectRow_b = selectRow(matrix, i + 1 + UBound(matrix, 1))
    End If
End Function
    Public Function p_selectRow_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_selectRow_b = make_funPointer(AddressOf selectRow_b, firstParam, secondParam)
    End Function

'特定列の取得（LBound基準）
'index < 0 の場合は後ろから取得
Public Function selectCol_b(ByRef matrix As Variant, ByRef j As Variant) As Variant
    If 0 <= j Then
        selectCol_b = selectCol(matrix, j + LBound(matrix, 2))
    Else
        selectCol_b = selectCol(matrix, j + 1 + UBound(matrix, 2))
    End If
End Function
    Public Function p_selectCol_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_selectCol_b = make_funPointer(AddressOf selectCol_b, firstParam, secondParam)
    End Function

'配列の特定行をデータで埋める（LBound基準）
Public Sub fillRow_b(ByRef matrix As Variant, ByVal i As Long, ByRef data As Variant)
    If 0 <= i Then
        Call fillRow(matrix, i + LBound(matrix, 1), data)
    Else
        Call fillRow(matrix, i + 1 + UBound(matrix, 1), data)
    End If
End Sub

'配列の特定行をデータで埋めてmoveして返す（LBound基準）
Public Function fillRow_b_move(ByRef matrix As Variant, ByVal i As Long, ByRef data As Variant) As Variant
    Call fillRow_b(matrix, i, data)
    fillRow_b_move = moveVariant(matrix)
End Function

'配列の特定列をデータで埋める（LBound基準）
Public Sub fillCol_b(ByRef matrix As Variant, ByVal j As Long, ByRef data As Variant)
    If 0 <= j Then
        Call fillCol(matrix, j + LBound(matrix, 2), data)
    Else
        Call fillCol(matrix, j + 1 + UBound(matrix, 2), data)
    End If
End Sub

'配列の特定列をデータで埋めてmoveして返す（LBound基準）
Public Function fillCol_b_move(ByRef matrix As Variant, ByVal j As Long, ByRef data As Variant) As Variant
    Call fillCol_b(matrix, j, data)
    fillCol_b_move = moveVariant(matrix)
End Function

'*********************************************************************************
' 1次元配列vecの隣接する要素間で2項操作opを行う
' 出力列の要素数は元の要素数 - 1   (LBound = 0)
Public Function adjacent_op(ByRef op As Variant, ByRef vec As Variant) As Variant
    If is_bindFun(op) Then
        adjacent_op = zipWith(op, vec, tailN(vec, sizeof(vec) - 1))
    End If
End Function
    Public Function p_adjacent_op(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_adjacent_op = make_funPointer(AddressOf adjacent_op, firstParam, secondParam)
    End Function

' 2次元配列の行ごとに関数適用
Public Sub rowWise_change(ByRef matrix As Variant, ByRef funcs As Variant)
    Dim i As Long
    For i = 0 To min_fun(rowSize(matrix), sizeof(funcs)) - 1 Step 1
        If is_bindFun(getNth_b(funcs, i)) Then
            Call fillRow_b(matrix, i, mapF(getNth_b(funcs, i), selectRow_b(matrix, i)))
        End If
    Next i
End Sub

' 2次元配列の行ごとに関数適用しmoveして返す
Public Function rowWise_change_move(ByRef matrix As Variant, ByRef funcs As Variant) As Variant
    Call rowWise_change(matrix, funcs)
    rowWise_change_move = moveVariant(matrix)
End Function
    Public Function p_rowWise_change_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_rowWise_change_move = make_funPointer(AddressOf rowWise_change_move, firstParam, secondParam)
    End Function

' 2次元配列の列ごとに関数適用
Public Sub columnWise_change(ByRef matrix As Variant, ByRef funcs As Variant)
    Dim j As Long
    For j = 0 To min_fun(colSize(matrix), sizeof(funcs)) - 1 Step 1
        If is_bindFun(getNth_b(funcs, j)) Then
            Call fillCol_b(matrix, j, mapF(getNth_b(funcs, j), selectCol_b(matrix, j)))
        End If
    Next j
End Sub

' 2次元配列の列ごとに関数適用しmoveして返す
Public Function colomnWise_change_move(ByRef matrix As Variant, ByRef funcs As Variant) As Variant
    Call columnWise_change(matrix, funcs)
    colomnWise_change_move = moveVariant(matrix)
End Function
    Public Function p_colomnWise_change_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_colomnWise_change_move = make_funPointer(AddressOf colomnWise_change_move, firstParam, secondParam)
    End Function

' 1次元配列の全要素の等値比較
Public Function equal_all(ByRef a As Variant, ByRef b As Variant) As Variant
    equal_all = equal_all_pred(p_equal, a, b)
End Function
    Public Function p_equal_all(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal_all = make_funPointer(AddressOf equal_all, firstParam, secondParam)
    End Function

' 1次元配列の全要素の等値比較（述語バージョン）
Public Function equal_all_pred(ByRef pred As Variant, ByRef a As Variant, ByRef b As Variant) As Variant
    If sizeof(a) = sizeof(b) Then
        equal_all_pred = IIf(sizeof(a) <= find_pred(p_equal(0), zipWith(pred, a, b)), _
                             1, _
                             0)
    Else
        equal_all_pred = 0
    End If
End Function

'******************************************************************************
' delimiterで区切られた文字列を関数列へマッピング
' strFuns   : 関数を表す文字列
' my_str2Fun: 文字列から関数へのマッピング関数
' delimiter : strFunsの区切り文字
' 例）%f%d%s%n → Array(f, d, s, n)
'******************************************************************************
Public Function splitStr2Funs(ByVal strFuns As String, _
                              ByRef my_str2Fun As Variant, _
                              ByVal delimiter As String) As Variant
    If left(strFuns, Len(delimiter)) = delimiter Then
        strFuns = right(strFuns, Len(strFuns) - Len(delimiter))
    End If
    splitStr2Funs = mapF(my_str2Fun, Split(strFuns, delimiter))
End Function

' （splitStr2Funs の対象関数）
' 文字列から集計関数へ変換
' 独自の変換関数を書くときはそのCase Else の中でこの関数を呼び出す形にするといいかも
' %t %tp  %top      : 先頭
' %b %btm %bottom   : 末尾
' %c %cnt %count    : 個数
' %s %sum %計       : 合計
' %a %avg %average  : 平均
' %max              : 最大
' %min              : 最少
Public Function str2SummaryFun(ByRef s As Variant, Optional ByRef other As Variant) As Variant
    Select Case StrConv(s, vbNarrow + vbLowerCase)
        Case "t", "tp", "top"
            str2SummaryFun = p_getNth_b(, 0)
        Case "b", "btm", "bottom"
            str2SummaryFun = p_getNth_b(, -1)
        Case "c", "cnt", "count"
            str2SummaryFun = p_sizeof()
        Case "s", "sum", "計"
            str2SummaryFun = p_foldl1(p_plus(yield_1, yield_2))
        Case "a", "avg", "average"
            str2SummaryFun = p_divide(p_foldl1(p_plus(yield_1, yield_2)), p_sizeof)
        Case "max"
            str2SummaryFun = p_foldl1(p_max(yield_1, yield_2))
        Case "min"
            str2SummaryFun = p_foldl1(p_min(yield_1, yield_2))
        Case ""
            str2SummaryFun = p_constant(other)
        Case Else
            str2SummaryFun = p_constant(s)
    End Select
End Function
    Public Function p_str2SummaryFun(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_str2SummaryFun = make_funPointer(AddressOf str2SummaryFun, firstParam, secondParam)
    End Function

' （splitStr2Funs の対象関数）
' 文字列から型変換関数へ変換
' 独自の変換関数を書くときはそのCase Else の中でこの関数を呼び出す形にするといいかも
' %s*  : Format( ,*)
' %d   : 整数化
' %f   : 実数化
' %s   : 文字列化
Public Function str2ConvertFun(ByRef s As Variant, ByRef dummy As Variant) As Variant
    Dim expr As String: expr = StrConv(s, vbNarrow + vbLowerCase)
    If left(expr, 1) = "s" Then
        If expr = "s" Then
            str2ConvertFun = p_CStr
        Else
            str2ConvertFun = p_format(, right(s, Len(s) - 1))
        End If
    Else
        Select Case expr
        Case "d"
            str2ConvertFun = p_CLng
        Case "f"
            str2ConvertFun = p_CDbl
        Case Else
            str2ConvertFun = Empty
        End Select
    End If
End Function
    Public Function p_str2ConvertFun(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_str2ConvertFun = make_funPointer(AddressOf str2ConvertFun, firstParam, secondParam)
    End Function

'******************************************************************************
' partition_points によるGROUP-BY
' matrix    : 対象配列（2次元配列またはジャグ配列）
' pp        : partition_points （集計する行範囲を区切る行番号の集合）
' strFuns   : 列ごとの集計関数を表す文字列
' my_str2Fun: 文字列から集計関数へのマッピング関数（str2SummaryFunがデフォルト）
' 例）group_by_partition_points(matrix, pp, "%t%c%s%a%min%max")
'******************************************************************************
Public Function group_by_partition_points(ByRef matrix As Variant, _
                                          ByRef pp As Variant, _
                                          ByRef strFuns As String, _
                                 Optional ByVal my_str2Fun As Variant) As Variant
    If IsMissing(my_str2Fun) Then my_str2Fun = p_str2SummaryFun(, "-")    'デフォルトの
    Dim funs As Variant
    funs = splitStr2Funs(strFuns, my_str2Fun, "%")
    Dim intervals As Variant
    intervals = adjacent_op(p_a__o, pp)
    Dim ranges As Variant
    ranges = mapF_swap(p_subM_R, matrix, intervals)
    group_by_partition_points = unzip(mapF_swap(p_summaryUnit, , funs, ranges), 2)
End Function
    ' 個々の集計行範囲の処理
    Private Function summaryUnit(ByRef matrix As Variant, ByRef funs As Variant) As Variant
        Select Case Dimension(matrix)
            Case 1: summaryUnit = zipWith(p_applyFun, unzip(matrix, 1), funs)
            Case 2: summaryUnit = zipWith(p_applyFun, zipR(matrix), funs)
        End Select
    End Function
    Private Function p_summaryUnit(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_summaryUnit = make_funPointer(AddressOf summaryUnit, firstParam, secondParam)
    End Function

'******************************************************************************
' csvファイルの1行を配列に分割
' expr      csvファイルの1行
' delim     区切り文字（省略時はカンマ）
'******************************************************************************
Function csv2Vector(ByRef expr As Variant, Optional ByRef delimiter As Variant) As Variant
    Dim delim As String
    delim = IIf(VarType(delimiter) = vbString, delimiter, ",")
    Dim line_s As String
    line_s = Replace(expr, """""", vbBack)      ' ""  -> vbBack Chr(8)
    Dim i As Long
    Dim inQuotationFlag As Boolean: inQuotationFlag = False
    For i = 1 To Len(line_s) Step 1
        If Mid(line_s, i, 1) = """" Then
            inQuotationFlag = Not inQuotationFlag
        End If
        If Mid(line_s, i, 1) = delim Then
            If Not inQuotationFlag Then
                Mid(line_s, i, 1) = Chr(0)
            End If
        End If
    Next i
    line_s = Replace(line_s, """", "")      ' " 消去
    line_s = Replace(line_s, vbBack, """")  ' vbBack -> "
    line_s = Replace(line_s, "\\t", vbLf)   ' \\t -> vbLf   Chr(10)
    line_s = Replace(line_s, "\t", vbTab)   ' \t  -> vbTab  Chr(9)
    line_s = Replace(line_s, vbLf, "\t")    ' vbLf   -> \\t
    csv2Vector = Split(line_s, Chr(0))
End Function
    Public Function p_csv2Vector(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_csv2Vector = make_funPointer_with_2nd_Default(AddressOf csv2Vector, firstParam, secondParam)
    End Function
