Attribute VB_Name = "Haskell_4_vector"
'Haskell_4_vector
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'====================================================================================================
' *_move系以外のPublicなFunctionは副作用なし
' 大多数のFunction に対して付随している p_Function は関数ポインタ。
'     mapF や zipWith の引数として使える。
'     p_Function : 裸の関数ポインタ、p_Function(a) : 第１引数を束縛、p_Function(, b) : 第２引数を束縛
'====================================================================================================
    ' Function  a_rows              全行番号の列挙
    ' Function  a_cols              全列番号の列挙
    ' Function  repeat              N個の値を並べる
    ' Function  iota                自然数の連続データ（正順・逆順）
    ' Function  a__a                自然数列 [from, to]
    ' Function  a__o                自然数列 [from, to)
    ' Function  o__a                自然数列 (from, to]
    ' Function  o__o                自然数列 (from, to)
    ' Function  headN               ベクトルの最初のN個
    ' Function  tailN               ベクトルの最後のN個
    ' Function  vector              スカラー、配列の1次元配列化
    ' Function  reverse             ベクトルを逆順に並べる
    ' Sub       rotate              1次元配列の回転
    ' Function  rotation            rotateした配列を返す
    ' Function  rotate_move         rotateしてmoveして返す
    ' Function  selectRow           特定行の取得
    ' Function  selectCol           特定列の取得
    ' Function  makeM               配列の作成
    ' Sub       fillM               配列をデータで埋める
    ' Function  fillM_move          fillMしてmoveして返す
    ' Sub       fillRow             配列の特定行をデータで埋める
    ' Function  fillRow_move        fillRowしてmoveして返す
    ' Sub       fillCol             配列の特定列をデータで埋める
    ' Function  fillCol_move        fillColしてmoveして返す
    ' Sub       fillPattern         1次元配列を他の1次元配列の繰り返しで埋める（回数指定可）
    ' Function  fillPattern_move    fillPatternしてmoveして返す
    ' Function  subV                1次元配列の部分配列を作成する
    ' Function  subV_if            　〃（範囲外のインデックスに対してEmptyが入る）
    ' Function  subM                配列の部分配列を作成する
    ' Function  subM_if            　〃（範囲外のインデックスに対してEmptyが入る）
    ' Function  filterR             ベクトル・配列の（行の）フィルタリング
    ' Function  filterC             ベクトル・配列の（列の）フィルタリング
    ' Function  catV                ベクトルを結合
    ' Function  catVs               ベクトルを結合（可変長引数）
    ' Function  catR                行方向に結合
    ' Function  catC                列方向に結合
    ' Function  transpose           配列の転置
    ' Function  zip                 ふたつの配列の対応する要素どうしをmakePairしてジャグ配列を作る
    ' Function  zipVs               複数の1次元配列をzip
    ' Function  unzip               'zipVsされたジャグ配列をほどいてzipVs前の1次元配列または2次元配列にする
    ' Function  zipR                2次元配列の各行ベクトルをzipVs
    ' Function  zipC                2次元配列の各列ベクトルをzipVs
    ' Function  makeSole            Array(a)作成
    ' Function  makePair            Array(a, b)作成
    ' Function  cons                配列の先頭に要素を追加
    ' Sub       push_back           1次元の配列の末尾に要素を追加（LBoundを0にする）
    ' Function  push_back_move      push_backしてmoveして返す
    ' Function  flatten             配列を平坦な1次元配列化する
    ' Function  product_set         ふたつのベクトルの直積に関数を適用した行列を作る
'====================================================================================================

'全行番号の列挙
Public Function a_rows(ByRef matrix As Variant, Optional ByRef dummy As Variant) As Variant
    a_rows = a__a(LBound(matrix, 1), UBound(matrix, 1))
End Function
    Public Function p_a_rows(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_a_rows = make_funPointer(AddressOf a_rows, firstParam, secondParam)
    End Function

'全列番号の列挙
Public Function a_cols(ByRef matrix As Variant, Optional ByRef dummy As Variant) As Variant
    If 2 <= Dimension(matrix) Then
        a_cols = iota(LBound(matrix, 2), UBound(matrix, 2))
    Else
        a_cols = VBA.Array()
    End If
End Function
    Public Function p_a_cols(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_a_cols = make_funPointer(AddressOf a_cols, firstParam, secondParam)
    End Function

'N個の値を並べる
Public Function repeat(ByRef v As Variant, ByRef n As Variant) As Variant
    If n < 1 Then
        repeat = VBA.Array()
    Else
        Dim i As Long
        Dim ret As Variant: ReDim ret(0 To n - 1)
        For i = 0 To n - 1 Step 1
            ret(i) = v
        Next i
        repeat = moveVariant(ret)
    End If
End Function
    Public Function p_repeat(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_repeat = make_funPointer(AddressOf repeat, firstParam, secondParam)
    End Function

' fromからtoまでの自然数を並べたベクトルを返す
' 両端入り。from <= to では昇順、from > to では逆順
Public Function iota(ByRef from_i As Variant, ByRef to_i As Variant) As Variant
    Dim i As Long, k As Long:   k = 0
    Dim ret As Variant:         ReDim ret(0 To VBA.Abs(to_i - from_i))
    Dim s_t_e_p As Long:        s_t_e_p = IIf(from_i < to_i, 1, -1)
    For i = from_i To to_i Step s_t_e_p
        ret(k) = i
        k = k + 1
    Next i
    iota = moveVariant(ret)
End Function
    Public Function p_iota(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_iota = make_funPointer(AddressOf iota, firstParam, secondParam)
    End Function

' 自然数列 [from, to]
Public Function a__a(ByRef from_i As Variant, ByRef to_i As Variant) As Variant
    If from_i <= to_i Then
        a__a = iota(from_i, to_i)
    Else
        a__a = VBA.Array()
    End If
End Function
    Public Function p_a__a(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_a__a = make_funPointer(AddressOf a__a, firstParam, secondParam)
    End Function

' 自然数列 [from, to)
Public Function a__o(ByRef from_i As Variant, ByRef to_i As Variant) As Variant
    If from_i < to_i Then
        Dim i As Long, k As Long:   k = 0
        Dim ret As Variant:         ReDim ret(0 To to_i - from_i - 1)
        For i = from_i To to_i - 1 Step 1
            ret(k) = i
            k = k + 1
        Next i
        a__o = moveVariant(ret)
    Else
        a__o = VBA.Array()
    End If
End Function
    Public Function p_a__o(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_a__o = make_funPointer(AddressOf a__o, firstParam, secondParam)
    End Function

' 自然数列 (from, to]
Public Function o__a(ByRef from_i As Variant, ByRef to_i As Variant) As Variant
    If from_i < to_i Then
        Dim i As Long, k As Long:   k = 0
        Dim ret As Variant:         ReDim ret(0 To to_i - from_i - 1)
        For i = from_i + 1 To to_i Step 1
            ret(k) = i
            k = k + 1
        Next i
        o__a = moveVariant(ret)
    Else
        o__a = VBA.Array()
    End If
End Function
    Public Function p_o__a(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_o__a = make_funPointer(AddressOf o__a, firstParam, secondParam)
    End Function

' 自然数列 (from, to)
Public Function o__o(ByRef from_i As Variant, ByRef to_i As Variant) As Variant
    If from_i + 1 < to_i Then
        Dim i As Long, k As Long:   k = 0
        Dim ret As Variant:         ReDim ret(0 To to_i - from_i - 2)
        For i = from_i + 1 To to_i - 1 Step 1
            ret(k) = i
            k = k + 1
        Next i
        o__o = moveVariant(ret)
    Else
        o__o = VBA.Array()
    End If
End Function
    Public Function p_o__o(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_o__o = make_funPointer(AddressOf o__o, firstParam, secondParam)
    End Function

'ベクトルの最初のN個
Public Function headN(ByRef vec As Variant, ByRef n As Variant) As Variant
    Dim lb As Long, i As Long
    Dim ret As Variant
    If n = 0 Then
        headN = VBA.Array()
    ElseIf n < 0 Then
        If 0 < sizeof(vec) + n Then
            headN = headN(vec, sizeof(vec) + n)
        Else
            headN = VBA.Array()
        End If
    ElseIf sizeof(vec) <= n Then
        ret = vec
        changeLBound ret, 0
        ReDim Preserve ret(0 To n - 1)
        swapVariant headN, ret
    Else
        lb = LBound(vec)
        ReDim ret(0 To n - 1)
        For i = 0 To n - 1 Step 1
            ret(i) = vec(i + lb)
        Next i
        headN = moveVariant(ret)
    End If
End Function
    Public Function p_headN(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_headN = make_funPointer(AddressOf headN, firstParam, secondParam)
    End Function

'ベクトルの最後のN個
Public Function tailN(ByRef vec As Variant, ByRef n As Variant) As Variant
    Dim lb As Long, i As Long
    Dim ret As Variant
    If n = 0 Then
        tailN = VBA.Array()
    ElseIf n < 0 Then
        If 0 < sizeof(vec) + n Then
            tailN = tailN(vec, sizeof(vec) + n)
        Else
            tailN = VBA.Array()
        End If
    ElseIf sizeof(vec) = n Then
        tailN = vec
        changeLBound tailN, 0
    ElseIf sizeof(vec) < n Then
        tailN = catV(makeM(n - sizeof(vec)), vec)
    Else
        lb = UBound(vec) - n + 1
        ReDim ret(0 To n - 1)
        For i = 0 To n - 1 Step 1
            ret(i) = vec(i + lb)
        Next i
        tailN = moveVariant(ret)
    End If
End Function
    Public Function p_tailN(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_tailN = make_funPointer(AddressOf tailN, firstParam, secondParam)
    End Function

'スカラー、配列のベクトル化
'デフォルトでC言語オーダー展開、orderがFalseに解釈できる場合はVBオーダー展開
Public Function vector(ByRef data As Variant, Optional ByRef order As Variant) As Variant
    Dim dimen As Long: dimen = Dimension(data)
    If dimen = 0 Then
        vector = VBA.Array(data)
        Exit Function
    ElseIf dimen = 1 Then
        vector = data
        changeLBound vector, 0
        Exit Function
    End If
    Dim c_order As Boolean: c_order = True
    If IsNumeric(order) And Not IsEmpty(order) Then c_order = order
    Dim i As Long, j As Long, counter As Long
    Dim ret As Variant, z As Variant
    ReDim ret(0 To sizeof(data) - 1)
    counter = 0
    If c_order Then
        If dimen = 2 Then
            For i = LBound(data, 1) To UBound(data, 1) Step 1
                For j = LBound(data, 2) To UBound(data, 2) Step 1
                    If IsObject(data(i, j)) Then
                        Set ret(counter) = data(i, j)
                    Else
                        ret(counter) = data(i, j)
                    End If
                    counter = counter + 1
                Next j
            Next i
        Else
            ReDim bound(1 To dimen, 0 To 1) As Long
            ReDim index(1 To dimen) As Long
            index(1) = -1
            For i = 1 To dimen Step 1
                bound(i, 0) = UBound(data, i) - LBound(data, i) + 1
            Next i
            bound(dimen, 1) = 1
            For i = dimen - 1 To 1 Step -1
                bound(i, 1) = bound(i + 1, 0) * bound(i + 1, 1)
            Next i
            For Each z In data
                counter = BoundaryControl(bound, index, 1)
                swapVariant ret(counter), z
            Next z
        End If
    Else
        For Each z In data
            swapVariant ret(counter), z
            counter = counter + 1
        Next z
    End If
    swapVariant vector, ret
End Function
    Public Function p_vector(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_vector = make_funPointer_with_2nd_Default(AddressOf vector, firstParam, secondParam)
    End Function

    Private Function BoundaryControl(ByRef bound As Variant, _
                                     ByRef index As Variant, _
                                     ByVal i As Long) As Long
        index(i) = index(i) + 1
        If index(i) = bound(i, 0) Then
            index(i) = 0
            BoundaryControl = BoundaryControl(bound, index, i + 1)
        Else
            BoundaryControl = 0
            Dim j As Long
            For j = LBound(index) To UBound(index) Step 1
                BoundaryControl = BoundaryControl + index(j) * bound(j, 1)
            Next j
        End If
    End Function

'ベクトルを逆順に並べる
Public Function reverse(ByRef vec As Variant) As Variant
    Dim ret As Variant
    Dim i As Long, j As Long
    If Dimension(vec) = 1 Then
        i = LBound(vec)
        j = UBound(vec)
        If VarType(vec) = VarType(Array()) Then
            ret = vec
            Do While i < j
                swapVariant ret(i), ret(j)
                i = i + 1
                j = j - 1
            Loop
        ElseIf IsObject(vec(LBound(vec))) Then
            ReDim ret(LBound(vec) To UBound(vec))
            Do While i <= j
                Set ret(i) = vec(j)
                If i <> j Then Set ret(j) = vec(i)
                i = i + 1
                j = j - 1
            Loop
        Else
            ret = vec
            Do While i < j
                ret(i) = vec(j)
                ret(j) = vec(i)
                i = i + 1
                j = j - 1
            Loop
        End If
    End If
    reverse = moveVariant(ret)
End Function

'1次元配列の回転
'[0,1,2,3,4,5] -> [1,2,3,4,5,0] (r=1)
'[0,1,2,3,4,5] -> [5,0,1,2,3,4] (r=-1)
Sub rotate(ByRef vec As Variant, ByVal shift As Long)
    If Dimension(vec) <> 1 Or sizeof(vec) = 0 Then Exit Sub
    If shift < 0 Then shift = (1 + (-shift) \ sizeof(vec)) * sizeof(vec) + shift
    shift = shift Mod sizeof(vec)
    If shift = 0 Then
        '
    ElseIf VarType(vec) = VarType(Array()) Then
        Call rotate_imple_V(vec, LBound(vec) + shift)
    ElseIf IsObject(vec(LBound(vec))) Then
        Call rotate_imple_O(vec, LBound(vec) + shift)
    Else
        Call rotate_imple_L(vec, LBound(vec) + shift)
    End If
End Sub

'1次元配列を回転した配列
Public Function rotation(ByRef vec As Variant, ByRef shift As Variant) As Variant
    Dim tmp As Variant
    tmp = vec
    rotation = rotate_move(tmp, shift)
End Function
    Public Function p_rotation(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_rotation = make_funPointer(AddressOf rotation, firstParam, secondParam)
    End Function

'rotationしてmoveして返す
Public Function rotate_move(ByRef vec As Variant, ByRef shift As Variant) As Variant
    rotate vec, shift
    rotate_move = moveVariant(vec)
End Function
    Function p_rotate_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_rotate_move = make_funPointer(AddressOf rotate_move, firstParam, secondParam)
    End Function
    '----------------
    Private Sub rotate_imple_V(ByRef vec As Variant, ByVal const_j As Long)
        Dim tmp As Variant:     Dim i As Long
        Dim k As Long:          k = LBound(vec)
        ReDim tmp(LBound(vec) To UBound(vec))
        For i = const_j To UBound(vec) Step 1
            swapVariant tmp(k), vec(i)
            k = k + 1
        Next i
        For i = LBound(vec) To const_j - 1 Step 1
            swapVariant tmp(k), vec(i)
            k = k + 1
        Next i
        vec = moveVariant(tmp)
    End Sub
    Private Sub rotate_imple_O(ByRef vec As Variant, ByVal const_j As Long)
        Dim tmp As Variant:     Dim i As Long
        Dim k As Long:          k = LBound(vec)
        ReDim tmp(LBound(vec) To UBound(vec))
        For i = const_j To UBound(vec) Step 1
            Set tmp(k) = vec(i)
            k = k + 1
        Next i
        For i = LBound(vec) To const_j - 1 Step 1
            Set tmp(k) = vec(i)
            k = k + 1
        Next i
        vec = moveVariant(tmp)
    End Sub
    Private Sub rotate_imple_L(ByRef vec As Variant, ByVal const_j As Long)
        Dim tmp As Variant:     Dim i As Long
        Dim k As Long:          k = LBound(vec)
        ReDim tmp(LBound(vec) To UBound(vec))
        For i = const_j To UBound(vec) Step 1
            tmp(k) = vec(i)
            k = k + 1
        Next i
        For i = LBound(vec) To const_j - 1 Step 1
            tmp(k) = vec(i)
            k = k + 1
        Next i
        vec = moveVariant(tmp)
    End Sub

'特定行の取得
Public Function selectRow(ByRef matrix As Variant, ByRef i As Variant) As Variant
    selectRow = VBA.Array()
    If Dimension(matrix) = 2 Then
        If LBound(matrix, 1) <= i And i <= UBound(matrix, 1) Then
            Dim j     As Long
            Dim ret   As Variant
            ReDim ret(LBound(matrix, 2) To UBound(matrix, 2))
            For j = LBound(matrix, 2) To UBound(matrix, 2) Step 1
                ret(j) = matrix(i, j)
            Next j
            selectRow = moveVariant(ret)
        End If
    End If
End Function
    Public Function p_selectRow(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_selectRow = make_funPointer(AddressOf selectRow, firstParam, secondParam)
    End Function

'特定列の取得
Public Function selectCol(ByRef matrix As Variant, ByRef j As Variant) As Variant
    selectCol = VBA.Array()
    If Dimension(matrix) = 2 Then
        If LBound(matrix, 2) <= j And j <= UBound(matrix, 2) Then
            Dim i     As Long
            Dim ret   As Variant
            ReDim ret(LBound(matrix, 1) To UBound(matrix, 1))
            For i = LBound(matrix, 1) To UBound(matrix, 1) Step 1
                ret(i) = matrix(i, j)
            Next i
            swapVariant selectCol, ret
        End If
    End If
End Function
    Public Function p_selectCol(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_selectCol = make_funPointer(AddressOf selectCol, firstParam, secondParam)
    End Function

'配列の作成      makeM(6, 3) => 6行(0,1,2,3,4,5) x 3列(0,1,2)
Public Function makeM(ByVal r As Long, Optional ByVal c As Variant, Optional ByRef data As Variant) As Variant
    Dim ret   As Variant
    ret = VBA.Array()
    If IsMissing(c) Then
        If 0 < r Then ReDim ret(0 To r - 1)
    Else
        If 0 < r And 0 < c Then ReDim ret(0 To r - 1, 0 To c - 1)
    End If
    If IsMissing(data) = False Then Call fillM(ret, data)
    makeM = moveVariant(ret)
End Function

'配列をデータで埋める
Public Sub fillM(ByRef matrix As Variant, ByRef data As Variant)
    Dim data_2  As Variant
    Dim stepN As Long: stepN = 1
    Dim swapFlag As Boolean: swapFlag = False
    If Dimension(data) = 0 Then
        data_2 = VBA.Array(data)
        stepN = 0
    ElseIf Dimension(data) = 1 Then
        swapVariant data_2, data
        swapFlag = True
    Else
        data_2 = vector(data)
    End If
    Dim i As Long, j As Long, k As Long
    k = LBound(data_2)
    Select Case Dimension(matrix)
    Case 1
        For i = LBound(matrix) To UBound(matrix) Step 1
            If UBound(data_2) < k Then Exit For
            matrix(i) = data_2(k)
            k = k + stepN
        Next i
    Case 2
        For i = LBound(matrix, 1) To UBound(matrix, 1) Step 1
            If UBound(data_2) < k Then Exit For
            For j = LBound(matrix, 2) To UBound(matrix, 2) Step 1
                If UBound(data_2) < k Then Exit For
                matrix(i, j) = data_2(k)
                k = k + stepN
            Next j
        Next i
    End Select
    If swapFlag Then swapVariant data_2, data
End Sub

'配列をデータで埋めてmoveして返す
Public Function fillM_move(ByRef matrix As Variant, ByRef data As Variant) As Variant
    Call fillM(matrix, data)
    fillM_move = moveVariant(matrix)
End Function

'配列の特定行をデータで埋める
Public Sub fillRow(ByRef matrix As Variant, ByVal i As Long, ByRef data As Variant)
    Dim j As Long, k As Long
    If Dimension(data) = 0 Then
        For j = LBound(matrix, 2) To UBound(matrix, 2) Step 1
            matrix(i, j) = data
        Next j
    ElseIf Dimension(data) = 1 Then
        k = LBound(data)
        For j = LBound(matrix, 2) To UBound(matrix, 2) Step 1
            If UBound(data) < k Then Exit For
            matrix(i, j) = data(k)
            k = k + 1
        Next j
    End If
End Sub

'配列の特定行をデータで埋めてmoveして返す
Public Function fillRow_move(ByRef matrix As Variant, ByVal i As Long, ByRef data As Variant) As Variant
    Call fillRow(matrix, i, data)
    fillRow_move = moveVariant(matrix)
End Function

    '((((配列の特定行をデータで埋める))))
    Private Sub fillRow_imple(ByRef matrix As Variant, _
                            ByVal i As Long, _
                        ByRef data As Variant, _
                    ByVal rrrr As Long)
        Dim j As Long, k As Long
        k = LBound(data, 2)
        For j = LBound(matrix, 2) To UBound(matrix, 2) Step 1
            matrix(i, j) = data(rrrr, k)
            k = k + 1
        Next j
    End Sub

'配列の特定列をデータで埋める
Public Sub fillCol(ByRef matrix As Variant, ByVal j As Long, ByRef data As Variant)
    Dim i As Long, k As Long
    If Dimension(data) = 0 Then
        For i = LBound(matrix, 1) To UBound(matrix, 1) Step 1
            matrix(i, j) = data
        Next i
    ElseIf Dimension(data) = 1 Then
        k = LBound(data)
        For i = LBound(matrix, 1) To UBound(matrix, 1) Step 1
            If UBound(data) < k Then Exit For
            matrix(i, j) = data(k)
            k = k + 1
        Next i
    End If
End Sub

'配列の特定列をデータで埋めてmoveして返す
Public Function fillCol_move(ByRef matrix As Variant, ByVal j As Long, ByRef data As Variant) As Variant
    Call fillCol(matrix, j, data)
    fillCol_move = moveVariant(matrix)
End Function
    
    '((((配列の特定列をデータで埋める))))
    Private Sub fillCol_imple(ByRef matrix As Variant, _
                            ByVal j As Long, _
                        ByRef data As Variant, _
                    ByVal cccc As Long)
        Dim i As Long, k As Long
        k = LBound(data, 1)
        For i = LBound(matrix, 1) To UBound(matrix, 1) Step 1
            matrix(i, j) = data(k, cccc)
            k = k + 1
        Next i
    End Sub

'1次元配列を他の1次元配列の繰り返しで埋める（回数指定可）
Sub fillPattern(ByRef vec As Variant, ByRef pattern As Variant, Optional ByVal counter As Long = -1)
    Dim ubm As Long:    ubm = UBound(vec)
    Dim ubp As Long:    ubp = UBound(pattern)
    Dim lbp As Long:    lbp = LBound(pattern)
    Dim i As Long:  i = LBound(vec)
    Dim k As Long:  k = LBound(pattern)
    Do While i <= ubm And counter <> 0
        vec(i) = pattern(k)
        i = i + 1
        k = k + 1
        If ubp < k Then
            k = lbp
            counter = counter - 1
        End If
    Loop
End Sub

'1次元配列を他の1次元配列の繰り返しで埋めてmoveして返す
Public Function fillPattern_move(ByRef vec As Variant, ByRef pattern As Variant, Optional ByVal counter As Long = -1) As Variant
    fillPattern vec, pattern, counter
    fillPattern_move = moveVariant(vec)
End Function

'1次元配列の部分配列を作成する
Public Function subV(ByRef vec As Variant, ByRef index As Variant) As Variant
    subV = mapF_swap(p_getNth, vec, , index)
    changeLBound subV, LBound(vec)
End Function
    Public Function p_subV(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subV = make_funPointer(AddressOf subV, firstParam, secondParam)
    End Function

'1次元配列の部分配列を作成する（範囲外のインデックスに対してEmptyが入る）
Public Function subV_if(ByRef vec As Variant, ByRef index As Variant) As Variant
    subV_if = mapF_swap(p_getNth_if, vec, , index)
    changeLBound subV_if, LBound(vec)
End Function
    Public Function p_subV_if(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subV_if = make_funPointer(AddressOf subV_if, firstParam, secondParam)
    End Function
    Private Function getNth_if(ByRef vec As Variant, ByRef index As Variant) As Variant
        If LBound(vec) <= index And index <= UBound(vec) Then
            getNth_if = vec(index)
        End If
    End Function
    Private Function p_getNth_if(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getNth_if = make_funPointer(AddressOf getNth_if, firstParam, secondParam)
    End Function

'配列の部分配列を作成する
Public Function subM(ByRef matrix As Variant, Optional ByRef rows As Variant, Optional ByRef cols As Variant) As Variant
    subM = subM_imple(matrix, False, rows, cols)
End Function

'配列の部分配列を作成する（範囲外のインデックスに対してEmptyが入る）
Public Function subM_if(ByRef matrix As Variant, Optional ByRef rows As Variant, Optional ByRef cols As Variant) As Variant
    subM_if = subM_imple(matrix, True, rows, cols)
End Function

    Private Function subM_imple(ByRef matrix As Variant, _
                                ByVal isif As Boolean, _
                                Optional ByRef rows As Variant, _
                                Optional ByRef cols As Variant) As Variant
        Dim ret As Variant
        Select Case Dimension(matrix)
        Case 0
            ret = matrix
        Case 1
            If isif Then
                ret = subV_if(matrix, rows)
            Else
                ret = subV(matrix, rows)
            End If
        Case 2
            If IsMissing(rows) Then
                If IsArray(rows) Then   ' 意図的に Array() を与えられたケース
                    subM_imple = VBA.Array()
                    Exit Function
                Else
                    rows = a_rows(matrix)
                End If
            End If
            If IsMissing(cols) Then
                If IsArray(cols) Then   ' 意図的に Array() を与えられたケース
                    subM_imple = VBA.Array()
                    Exit Function
                Else
                    cols = a_cols(matrix)
                End If
            End If
            Dim i As Long, j As Long, counterR As Long, counterC As Long
            counterR = LBound(matrix, 1)
            counterC = LBound(matrix, 2)
            If 0 < sizeof(rows) And 0 < sizeof(cols) Then
                ReDim ret(counterR To counterR - 1 + sizeof(rows), counterC To counterC - 1 + sizeof(cols))
            End If
            If isif Then
                For i = LBound(rows) To UBound(rows) Step 1
                    counterC = LBound(matrix, 2)
                    If LBound(matrix, 1) <= rows(i) And rows(i) <= UBound(matrix, 1) Then
                        For j = LBound(cols) To UBound(cols) Step 1
                            If LBound(matrix, 2) <= cols(j) And cols(j) <= UBound(matrix, 2) Then
                                ret(counterR, counterC) = matrix(rows(i), cols(j))
                            End If
                            counterC = counterC + 1
                        Next j
                    End If
                    counterR = counterR + 1
                Next i
            Else
                For i = LBound(rows) To UBound(rows) Step 1
                    counterC = LBound(matrix, 2)
                    For j = LBound(cols) To UBound(cols) Step 1
                        ret(counterR, counterC) = matrix(rows(i), cols(j))
                        counterC = counterC + 1
                    Next j
                    counterR = counterR + 1
                Next i
            End If
        End Select
        subM_imple = moveVariant(ret)
    End Function

'ベクトル・配列の（行の）フィルタリング
'Flgは 0/1
Public Function filterR(ByRef data As Variant, ByRef flg As Variant) As Variant
    Dim indice As Variant
    Dim i As Long, j As Long, counter As Long, NN As Long
    NN = min_fun(sizeof(flg), rowSize(data))
    indice = makeM(NN)
    i = LBound(data, 1)
    counter = -1
    For j = LBound(flg) To LBound(flg) + NN - 1 Step 1
        If flg(j) <> 0 Then
            counter = counter + 1
            indice(counter) = i
        End If
        i = i + 1
    Next j
    If 0 <= counter Then
        ReDim Preserve indice(0 To counter)
    Else
        indice = Array()
    End If
    filterR = subM(data, indice)
End Function
    Public Function p_filterR(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_filterR = make_funPointer(AddressOf filterR, firstParam, secondParam)
    End Function

'ベクトル・配列の（列の）フィルタリング
'Flgは 0/1
Public Function filterC(ByRef data As Variant, ByRef flg As Variant) As Variant
    If Dimension(data) = 2 Then
        Dim indice As Variant
        Dim i As Long, j As Long, counter As Long, NN As Long
        NN = min_fun(sizeof(flg), colSize(data))
        indice = makeM(NN)
        i = LBound(data, 2)
        counter = -1
        For j = LBound(flg) To LBound(flg) + NN - 1 Step 1
            If flg(j) <> 0 Then
                counter = counter + 1
                indice(counter) = i
            End If
            i = i + 1
        Next j
        If 0 <= counter Then
            ReDim Preserve indice(0 To counter)
        Else
            indice = Array()
        End If
        filterC = subM(data, , indice)
    Else
        filterC = VBA.Array()
    End If
End Function
    Public Function p_filterC(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_filterC = make_funPointer(AddressOf filterC, firstParam, secondParam)
    End Function

'ベクトルを結合
Public Function catV(ByRef v1 As Variant, ByRef v2 As Variant) As Variant
    Dim i As Long, counter As Long
    Dim ret As Variant
    If Dimension(v1) = 1 And Dimension(v2) = 1 Then
        If rowSize(v1) = 0 Then
            ret = v2
        ElseIf rowSize(v2) = 0 Then
            ret = v1
        Else
            ret = v1
            ReDim Preserve ret(0 To rowSize(v1) + rowSize(v2) - 1)
            counter = rowSize(v1)
            For i = LBound(v2) To UBound(v2) Step 1
                ret(counter) = v2(i)
                counter = counter + 1
            Next i
        End If
        swapVariant catV, ret
    ElseIf Dimension(v1) <> 1 And Dimension(v2) = 1 Then
        catV = catV(vector(v1), v2)
    ElseIf Dimension(v1) = 1 And Dimension(v2) <> 1 Then
        catV = catV(v1, vector(v2))
    Else
        catV = catV(vector(v1), vector(v2))
    End If
End Function
    Public Function p_catV(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_catV = make_funPointer(AddressOf catV, firstParam, secondParam)
    End Function

'ベクトルを結合（可変長引数）
Public Function catVs(ParamArray vectors() As Variant) As Variant
    Dim i As Long
    Dim tmp As Variant
    If LBound(vectors) <= UBound(vectors) Then
        ReDim tmp(LBound(vectors) To UBound(vectors))
        For i = LBound(vectors) To UBound(vectors)
            swapVariant vectors(i), tmp(i)
        Next i
        catVs = foldl1(p_catV, tmp)
        For i = LBound(vectors) To UBound(vectors)
            swapVariant vectors(i), tmp(i)
        Next i
    End If
End Function

'行方向に結合
Public Function catR(ByRef matrix1 As Variant, ByRef matrix2 As Variant) As Variant
    If rowSize(matrix1) = 0 Then        ' 非配列または空の配列
        If rowSize(matrix2) = 0 Then        ' 非配列または空の配列
            catR = VBA.Array()
        ElseIf Dimension(matrix2) = 1 Then
            catR = makeM(1, rowSize(matrix2), matrix2)
        Else
            catR = matrix2
        End If
    ElseIf Dimension(matrix1) = 1 Then
        catR = catR(makeM(1, rowSize(matrix1), matrix1), matrix2)
    Else
        If rowSize(matrix2) = 0 Then        ' 非配列または空の配列
            catR = matrix1
        ElseIf Dimension(matrix2) = 1 Then
            catR = catR(matrix1, makeM(1, rowSize(matrix2), matrix2))
        ElseIf colSize(matrix1) <> colSize(matrix2) Then
            catR = VBA.Array()
        Else
            Dim i As Long, counter As Long
            Dim ret As Variant
            ret = makeM(rowSize(matrix1) + rowSize(matrix2), colSize(matrix1))
            counter = 0
            For i = LBound(matrix1, 1) To UBound(matrix1, 1) Step 1
                Call fillRow_imple(ret, counter, matrix1, i)
                counter = counter + 1
            Next i
            For i = LBound(matrix2, 1) To UBound(matrix2, 1) Step 1
                Call fillRow_imple(ret, counter, matrix2, i)
                counter = counter + 1
            Next i
            swapVariant catR, ret
        End If
    End If
End Function
    Public Function p_catR(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_catR = make_funPointer(AddressOf catR, firstParam, secondParam)
    End Function

'列方向に結合
Public Function catC(ByRef matrix1 As Variant, ByRef matrix2 As Variant) As Variant
    If rowSize(matrix1) = 0 Then        ' 非配列または空の配列
        If rowSize(matrix2) = 0 Then        ' 非配列または空の配列
            catC = VBA.Array()
        ElseIf Dimension(matrix2) = 1 Then
            catC = makeM(rowSize(matrix2), 1, matrix2)
        Else
            catC = matrix2
        End If
    ElseIf Dimension(matrix1) = 1 Then
        catC = catC(makeM(rowSize(matrix1), 1, matrix1), matrix2)
    Else
        If rowSize(matrix2) = 0 Then        ' 非配列または空の配列
            catC = matrix1
        ElseIf Dimension(matrix2) = 1 Then
            catC = catC(matrix1, makeM(rowSize(matrix2), 1, matrix2))
        ElseIf rowSize(matrix1) <> rowSize(matrix2) Then
            catC = VBA.Array()
        Else
            Dim j As Long, counter As Long
            Dim ret As Variant
            ret = makeM(rowSize(matrix1), colSize(matrix1) + colSize(matrix2))
            counter = 0
            For j = LBound(matrix1, 2) To UBound(matrix1, 2) Step 1
                Call fillCol_imple(ret, counter, matrix1, j)
                counter = counter + 1
            Next j
            For j = LBound(matrix2, 2) To UBound(matrix2, 2) Step 1
                Call fillCol_imple(ret, counter, matrix2, j)
                counter = counter + 1
            Next j
            swapVariant catC, ret
        End If
    End If
End Function
    Public Function p_catC(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_catC = make_funPointer(AddressOf catC, firstParam, secondParam)
    End Function

'配列の行/列の転置
Public Function transpose(ByRef matrix As Variant) As Variant
    Dim i As Long, j As Long
    Dim ret As Variant
    Select Case Dimension(matrix)
    Case 0
        transpose = matrix
    Case 1
        If rowSize(matrix) = 0 Then
            transpose = VBA.Array()
        Else
            transpose = makeM(sizeof(matrix), 1, matrix)
        End If
    Case 2
        ReDim ret(LBound(matrix, 2) To UBound(matrix, 2), LBound(matrix, 1) To UBound(matrix, 1))
        For i = LBound(matrix, 2) To UBound(matrix, 2)
            For j = LBound(matrix, 1) To UBound(matrix, 1)
                ret(i, j) = matrix(j, i)
            Next j
        Next i
        swapVariant transpose, ret
    Case Else
        transpose = VBA.Array()
    End Select
End Function

'ふたつの配列の対応する要素どうしをmakePairしてジャグ配列を作る
Public Function zip(ByRef a As Variant, ByRef b As Variant) As Variant
    zip = zipWith(p_makePair, a, b)
End Function
    Public Function p_zip(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_zip = make_funPointer(AddressOf zip, firstParam, secondParam)
    End Function

    ' zipVsのサブルーチン
    Private Function zipVs_imple(ByRef m As Variant, ByRef a As Variant) As Variant
        Dim i As Long, j As Long: j = m(0)
        Dim k As Long: k = 0
        For i = LBound(a) To UBound(a) Step 1
            m(1)(k)(j) = a(i)
            k = k + 1
        Next i
        m(0) = m(0) + 1
        swapVariant zipVs_imple, m
    End Function
        Public Function p_zipVs_imple(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
            p_zipVs_imple = make_funPointer(AddressOf zipVs_imple, firstParam, secondParam)
        End Function

'複数の1次元配列をzip
Public Function zipVs(ByRef vectors As Variant) As Variant
    Dim ret As Variant
    ret = VBA.Array(0, repeat(makeM(sizeof(vectors)), sizeof(vectors(LBound(vectors)))))
    ret = foldl(p_zipVs_imple, ret, vectors)
    swapVariant zipVs, ret(1)
End Function

'２次元配列の各行ベクトルをzipVs
Public Function zipR(ByRef m As Variant, Optional ByRef target As Variant) As Variant
    If IsMissing(target) Then
        Dim ret As Variant
        ReDim ret(LBound(m, 2) To UBound(m, 2))
        Dim i As Long
        For i = LBound(m, 2) To UBound(m, 2) Step 1
            ret(i) = selectCol(m, i)
        Next i
        swapVariant zipR, ret
    Else
        zipR = zipR(subM(m, target))
    End If
End Function

'２次元配列の各列ベクトルをzipVs
Public Function zipC(ByRef m As Variant, Optional ByRef target As Variant) As Variant
    If IsMissing(target) Then
        Dim ret As Variant
        ReDim ret(LBound(m, 1) To UBound(m, 1))
        Dim i As Long
        For i = LBound(m, 1) To UBound(m, 1) Step 1
            ret(i) = selectRow(m, i)
        Next i
        swapVariant zipC, ret
    Else
        zipC = zipC(subM(m, , target))
    End If
End Function

'zipVsされたジャグ配列をほどいてzipVs前の1次元配列または2次元配列にする
Public Function unzip(ByRef vec As Variant, Optional ByVal dimen As Long = 1) As Variant
    Dim colLen As Long, i As Long, j As Long, counter As Long
    Dim ret As Variant, z As Variant
    unzip = VBA.Array()
    colLen = 0
    For counter = LBound(vec) To UBound(vec) Step 1
        If colLen < sizeof(vec(counter)) Then colLen = sizeof(vec(counter))
    Next counter
    If colLen = 0 Then Exit Function
    If dimen = 1 Then
        ReDim ret(0 To colLen - 1)
        For j = LBound(ret) To UBound(ret) Step 1
            ReDim z(0 To sizeof(vec) - 1)
            counter = 0
            For i = LBound(vec) To UBound(vec) Step 1
                If j <= UBound(vec(i)) Then z(counter) = vec(i)(j)
                counter = counter + 1
            Next i
            swapVariant ret(j), z
        Next j
    Else
        ReDim ret(0 To sizeof(vec) - 1, 0 To colLen - 1)
        counter = LBound(vec)
        For i = LBound(ret, 1) To UBound(ret, 1) Step 1
            Call fillRow(ret, i, vec(counter))
            counter = counter + 1
        Next i
    End If
    unzip = moveVariant(ret)
End Function

' Array(a)作成
Public Function makeSole(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    makeSole = VBA.Array(a)
End Function
    Public Function p_makeSole(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_makeSole = make_funPointer(AddressOf makeSole, firstParam, secondParam)
    End Function

' Array(a, b)作成
Public Function makePair(ByRef a As Variant, ByRef b As Variant) As Variant
    makePair = VBA.Array(a, b)
End Function
    Public Function p_makePair(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_makePair = make_funPointer(AddressOf makePair, firstParam, secondParam)
    End Function

' 配列の先頭に要素を追加
Public Function cons(ByRef a As Variant, ByRef vec As Variant) As Variant
    cons = catV(Array(a), vec)
End Function
    Public Function p_cons(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_cons = make_funPointer(AddressOf cons, firstParam, secondParam)
    End Function

' １次元配列の末尾に要素を追加（LBoundを0にする）
Public Sub push_back(ByRef vec As Variant, ByRef a As Variant)
    If Dimension(vec) = 1 Then
        changeLBound vec, 0
        If UBound(vec) < 0 Then
            ReDim vec(0 To 0)
        Else
            ReDim Preserve vec(0 To UBound(vec) + 1)
        End If
        vec(UBound(vec)) = a
    End If
End Sub

' push_backしてmoveして返す
Public Function push_back_move(ByRef vec As Variant, ByRef a As Variant) As Variant
    push_back vec, a
    swapVariant push_back_move, vec
End Function
    Public Function p_push_back_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_push_back_move = make_funPointer(AddressOf push_back_move, firstParam, secondParam)
    End Function

' 配列を平坦な1次元配列化する
Public Function flatten(ByRef vec As Variant, Optional ByRef dummy As Variant) As Variant
    flatten = VBA.Array()
    flatten_imple flatten, vec
End Function
    Public Function p_flatten(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_flatten = make_funPointer(AddressOf flatten, firstParam, secondParam)
    End Function
    
    Private Sub flatten_imple(ByRef vec As Variant, ByRef x As Variant)
        Select Case Dimension(x)
        Case 0
            push_back vec, x
        Case 1
            Dim i As Long
            For i = LBound(x) To UBound(x) Step 1
                flatten_imple vec, x(i)    ' 再帰
            Next i
        Case 2
            Dim j As Long
            For i = LBound(x, 1) To UBound(x, 1) Step 1
                For j = LBound(x, 2) To UBound(x, 2) Step 1
                    flatten_imple vec, x(i, j)    ' 再帰
                Next j
            Next i
        Case Else
            flatten_imple vec, vector(x)
        End Select
    End Sub

'ベクトルの直積に関数を適用した行列を作る
Public Function product_set(ByRef pCallback As Variant, ByRef a As Variant, ByRef b As Variant) As Variant
    Dim z As Variant, k As Long
    Dim ret As Variant:     ReDim ret(LBound(a) To UBound(a), LBound(b) To UBound(b))
    If rowSize(a) < rowSize(b) Then
        k = LBound(a)
        For Each z In a
            Call fillRow(ret, k, mapF(bind1st(pCallback, z), b))
            k = k + 1
        Next z
    Else
        k = LBound(b)
        For Each z In b
            Call fillCol(ret, k, mapF(bind2nd(pCallback, z), a))
            k = k + 1
        Next z
    End If
    product_set = moveVariant(ret)
End Function
