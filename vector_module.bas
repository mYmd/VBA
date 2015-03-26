Attribute VB_Name = "vector_module"
'vector_module
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'====================================================================================================
' すべてのPublicなFunctionは副作用なし
' 大多数のFunction に対して付随している p_Function は関数ポインタ。
'     mapF や zipWith の引数として使える。
'     p_Function : 裸の関数ポインタ、p_Function(a) : 第１引数を束縛、p_Function(, b) : 第２引数を束縛
'====================================================================================================
    ' Function  rowSize             配列の行数
    ' Function  colSize             配列の列数
    ' Function  sizeof              配列の全要素数
    ' Function  a_rows              全行番号の列挙
    ' Function  a_cols              全列番号の列挙
    ' Function  repeat              N個の値を並べる
    ' Function  iota                自然数の連続データまたは同一のスカラーを繰り返したベクトルを返す
    ' Function  headN               ベクトルの最初のN個をとる
    ' Function  tailN               ベクトルの最後のN個をとる
    ' Function  vector              スカラー、配列のベクトル化
    ' Function  reverse             ベクトルを逆順に並べる
    ' Function  selectRow           特定行の取得
    ' Function  selectCol           特定列の取得
    ' Function  makeM               配列の作成
    ' Sub       fillM               配列をデータで埋める
    ' Sub       fillRow             配列の特定行をデータで埋める
    ' Sub       fillCol             配列の特定列をデータで埋める
    ' Function  subM                配列の部分配列を作成する
    ' Function  filterR             ベクトル・配列の（行の）取捨をする
    ' Function  filterC             ベクトル・配列の（列の）取捨をする
    ' Function  catV                ベクトルを結合
    ' Function  catVs               ベクトルを結合（可変長引数）
    ' Function  catR                行方向に結合
    ' Function  catC                列方向に結合
    ' Function  transpose           配列の転置
    ' Function  zip                 ふたつの配列の対応する要素どうしをcatV(ベクトル結合)してジャグ配列を作る
    ' Function  zipVs               可変長引数zip
    ' Function  unzip               zipされたジャグ配列をほどいて複数の1次元配列または一つの2次元配列に展開する
    ' Function  makePair            Array(a, b)作成
    ' Function  product_set         ふたつのベクトルの直積に関数を適用した行列を作る
'====================================================================================================

'配列の行数
Public Function rowSize(ByRef data As Variant) As Long
    Select Case Dimension(data)
    Case 0
        rowSize = 0
    Case Else
        rowSize = 1 + UBound(data) - LBound(data)
    End Select
End Function

'配列の列数
Public Function colSize(ByRef data As Variant) As Long
    Select Case Dimension(data)
    Case 0, 1
        colSize = 0
    Case Else
        colSize = 1 + UBound(data, 2) - LBound(data, 2)
    End Select
End Function

'配列の全要素数
Public Function sizeof(ByRef data As Variant) As Long
    Dim i As Long, d As Long
    
    d = Dimension(data)
    sizeof = 1
    For i = 1 To d Step 1:        sizeof = sizeof * (1 + UBound(data, i) - LBound(data, i)):    Next i
End Function

'全行番号の列挙
Public Function a_rows(ByRef matrix As Variant) As Variant
    a_rows = iota(LBound(matrix, 1), UBound(matrix, 1))
End Function

'全列番号の列挙
Public Function a_cols(ByRef matrix As Variant) As Variant
    a_cols = iota(LBound(matrix, 2), UBound(matrix, 2))
End Function

'N個の値を並べる
Public Function repeat(ByRef v As Variant, ByVal N As Long) As Variant
    Dim ret As Variant
    Dim i As Long
    
    If N < 1 Then repeat = VBA.Array(): Exit Function
    ReDim ret(0 To N - 1)
    For i = 0 To N - 1 Step 1:         ret(i) = v:       Next i
    repeat = ret
End Function

'from_iからto_iまでの自然数を並べたベクトルを返す
Public Function iota(ByVal from_i As Long, ByVal to_i As Long) As Variant
    Dim ret   As Variant
    Dim i As Long, k As Long, s_t_e_p As Long
    
    ReDim ret(0 To IIf(from_i < to_i, to_i - from_i, from_i - to_i))
    s_t_e_p = IIf(from_i < to_i, 1, -1)
    k = 0
    For i = from_i To to_i Step s_t_e_p
        ret(k) = i
        k = k + 1
    Next i
    iota = ret
End Function

'ベクトルの最初のN個をとる
Public Function headN(ByRef vec As Variant, ByRef N As Variant) As Variant
    Dim lb As Long, i As Long
    Dim ret As Variant
    
    If N < 1 Then
        headN = VBA.Array()
    ElseIf sizeof(vec) < N Then
        headN = vec
    Else
        lb = LBound(vec)
        ReDim ret(0 To N - 1)
        For i = 0 To N - 1 Step 1
            ret(i) = vec(i + lb)
        Next i
        headN = ret
    End If
End Function
    Public Function p_headN(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_headN = make_funPointer(AddressOf headN, firstParam, secondParam)
    End Function

'ベクトルの最後のN個をとる
Public Function tailN(ByRef vec As Variant, ByRef N As Variant) As Variant
    Dim lb As Long, i As Long
    Dim ret As Variant
    
    If N < 1 Then
        tailN = VBA.Array()
    ElseIf sizeof(vec) < N Then
        tailN = vec
    Else
        lb = UBound(vec) - N + 1
        ReDim ret(0 To N - 1)
        For i = 0 To N - 1 Step 1
            ret(i) = vec(i + lb)
        Next i
        tailN = ret
    End If
End Function
    Public Function p_tailN(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_tailN = make_funPointer(AddressOf tailN, firstParam, secondParam)
    End Function

'スカラー、配列のベクトル化(行→列)
Public Function vector(data As Variant) As Variant
    Dim i As Long, j As Long, counter As Long
    Dim ret   As Variant
    
    Select Case Dimension(data)
    Case 0
        vector = VBA.Array(data)
    Case 1
        vector = data
    Case 2
        counter = 0
        ReDim ret(0 To sizeof(data) - 1)
        For i = LBound(data, 1) To UBound(data, 1) Step 1
            For j = LBound(data, 2) To UBound(data, 2) Step 1
                ret(counter) = data(i, j)
                counter = counter + 1
            Next j
        Next i
        vector = ret
    End Select
End Function

'ベクトルを逆順に並べる
Public Function reverse(ByRef data As Variant) As Variant
    Dim ret As Variant
    Dim i  As Long, j As Long

    ret = data
    If Dimension(data) = 1 Then
         i = LBound(data)
         j = UBound(data)
        Do While i < j
            ret(i) = data(j)
            ret(j) = data(i)
            i = i + 1
            j = j - 1
        Loop
    End If
    reverse = ret
End Function

'特定行の取得
Public Function selectRow(data As Variant, ByRef i As Variant) As Variant
    Dim j     As Long
    Dim ret   As Variant

    If i < LBound(data, 1) Or UBound(data, 1) < i Then
        selectRow = VBA.Array()
    Else
        ReDim ret(LBound(data, 2) To UBound(data, 2))
        For j = LBound(data, 2) To UBound(data, 2) Step 1
            ret(j) = data(i, j)
        Next j
        selectRow = ret
    End If
End Function
    Public Function p_selectRow(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_selectRow = make_funPointer(AddressOf selectRow, firstParam, secondParam)
    End Function

'特定列の取得
Public Function selectCol(data As Variant, ByRef j As Variant) As Variant
    Dim i     As Long
    Dim ret   As Variant

    If j < LBound(data, 2) Or UBound(data, 2) < j Then
        selectCol = VBA.Array()
    Else
        ReDim ret(LBound(data, 1) To UBound(data, 1))
        For i = LBound(data, 1) To UBound(data, 1) Step 1
            ret(i) = data(i, j)
        Next i
        selectCol = ret
    End If
End Function
    Public Function p_selectCol(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_selectCol = make_funPointer(AddressOf selectCol, firstParam, secondParam)
    End Function

'配列の作成      makeM(6, 3) => 6行(0,1,2,3,4,5) x 3列(0,1,2)
Public Function makeM(ByVal r As Long, Optional ByVal c As Variant, Optional ByRef data As Variant) As Variant
    Dim ret   As Variant
    
    If IsMissing(c) Then
        ReDim ret(0 To r - 1)
    Else
        ReDim ret(0 To r - 1, 0 To c - 1)
    End If
    If IsMissing(data) = False Then Call fillM(ret, data)
    makeM = ret
End Function

'配列をデータで埋める
Public Sub fillM(ByRef matrix As Variant, ByRef data As Variant)
    Dim i    As Long, j As Long, k As Long
    Dim data_2  As Variant

    If Dimension(data) = 0 Then
        data_2 = repeat(data, sizeof(matrix))
    Else
        data_2 = vector(data)
    End If
    k = LBound(data_2)
    Select Case Dimension(matrix)
    Case 1
        For i = LBound(matrix) To UBound(matrix) Step 1
            matrix(i) = data_2(k)
            k = k + 1
            If UBound(data_2) < k Then Exit For
        Next i
    Case 2
        For i = LBound(matrix, 1) To UBound(matrix, 1) Step 1
            If UBound(data_2) < k Then Exit For
            For j = LBound(matrix, 2) To UBound(matrix, 2) Step 1
                matrix(i, j) = data_2(k)
                k = k + 1
                If UBound(data_2) < k Then Exit For
            Next j
        Next i
    End Select
End Sub

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
            matrix(i, j) = data(k)
            k = k + 1
            If UBound(data) < k Then Exit For
        Next j
    End If
End Sub

'((((配列の特定行をデータで埋める))))
Private Sub fillRow_imple(ByRef matrix As Variant, ByVal i As Long, ByRef data As Variant, ByVal rrrr As Long)
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
            matrix(i, j) = data(k)
            k = k + 1
            If UBound(data) < k Then Exit For
        Next i
    End If
End Sub

'((((配列の特定列をデータで埋める))))
Private Sub fillCol_imple(ByRef matrix As Variant, ByVal j As Long, ByRef data As Variant, ByVal cccc As Long)
    Dim i As Long, k As Long
    
    k = LBound(data, 1)
    For i = LBound(matrix, 1) To UBound(matrix, 1) Step 1
        matrix(i, j) = data(k, cccc)
        k = k + 1
    Next i
End Sub

'配列の部分配列を作成する
Public Function subM(matrix As Variant, Optional ByRef rows As Variant, Optional ByRef cols As Variant) As Variant
    Dim i As Long, j As Long, counterR As Long, counterC As Long
    Dim ret As Variant

    Select Case Dimension(matrix)
    Case 0
        subM = matrix
        Exit Function
    Case 1
        counterR = LBound(matrix, 1)
        ReDim ret(counterR To counterR - 1 + sizeof(rows))
        For i = LBound(rows) To UBound(rows) Step 1
            ret(counterR) = matrix(rows(i))
            counterR = counterR + 1
        Next i
    Case 2
        If IsMissing(rows) Then
            If IsArray(rows) Then
                subM = VBA.Array()
                Exit Function
            Else
                rows = a_rows(matrix)
            End If
        End If
        If IsMissing(cols) Then
            If IsArray(cols) Then
                subM = VBA.Array()
                Exit Function
            Else
                cols = a_cols(matrix)
            End If
        End If
        counterR = LBound(matrix, 1)
        counterC = LBound(matrix, 2)
        ReDim ret(counterR To counterR - 1 + sizeof(rows), counterC To counterC - 1 + sizeof(cols))
        For i = LBound(rows) To UBound(rows) Step 1
            counterC = LBound(matrix, 2)
            For j = LBound(cols) To UBound(cols) Step 1
                ret(counterR, counterC) = matrix(rows(i), cols(j))
                counterC = counterC + 1
            Next j
            counterR = counterR + 1
        Next i
    End Select
    subM = ret
End Function

    Private Function filter_imple(ByRef pos As Variant, ByRef index As Variant) As Variant
        If 0 <> index Then
            pos(2)(pos(3)) = pos(0) + pos(1)
            filter_imple = VBA.Array(pos(0) + 1, pos(1), pos(2), pos(3) + 1)
        Else
            filter_imple = VBA.Array(pos(0) + 1, pos(1), pos(2), pos(3))
        End If
    End Function


'ベクトル・配列の（行の）取捨をする
'Flgは 0/1
Public Function filterR(ByRef data As Variant, ByRef flg As Variant) As Variant
    Dim filterSize As Long, dataSize As Long
    Dim indice As Variant, tmpFlag As Variant
    
    filterSize = sizeof(flg)
    dataSize = rowSize(data)
    If dataSize < filterSize Then filterSize = dataSize
    tmpFlag = headN(flg, filterSize)
    ReDim indice(0 To -1 + count_if(p_notEqual(, 0), tmpFlag))
    indice = foldl(AddressOf filter_imple, VBA.Array(0, LBound(flg), indice, 0), tmpFlag)
    filterR = subM(data, indice(2))
End Function
    Public Function p_filterR(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_filterR = make_funPointer(AddressOf filterR, firstParam, secondParam)
    End Function

'ベクトル・配列の（列の）取捨をする
'Flgは 0/1
Public Function filterC(ByRef data As Variant, ByRef flg As Variant) As Variant
    Dim filterSize As Long, dataSize As Long
    Dim indice As Variant, tmpFlag As Variant
    
    filterSize = sizeof(flg)
    dataSize = colSize(data)
    If dataSize < filterSize Then filterSize = dataSize
    tmpFlag = headN(flg, filterSize)
    ReDim indice(0 To -1 + count_if(p_notEqual(, 0), tmpFlag))
    indice = foldl(AddressOf filter_imple, VBA.Array(0, LBound(flg), indice, 0), tmpFlag)
    filterC = subM(data, , indice(2))
End Function
    Public Function p_filterC(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_filterC = make_funPointer(AddressOf filterC, firstParam, secondParam)
    End Function

'ベクトルを結合
Function catV(ByRef v1 As Variant, ByRef v2 As Variant) As Variant
    Dim i As Long, counter As Long
    Dim ret As Variant
    
    If Dimension(v1) = 1 And Dimension(v2) = 1 Then
        ReDim ret(0 To sizeof(v1) + sizeof(v2) - 1)
        counter = 0
        For i = LBound(v1) To UBound(v1) Step 1
            ret(counter) = v1(i)
            counter = counter + 1
        Next i
        For i = LBound(v2) To UBound(v2) Step 1
            ret(counter) = v2(i)
            counter = counter + 1
        Next i
        catV = ret
    ElseIf Dimension(v1) <> 1 And Dimension(v2) = 1 Then
        catV = catV(vector(v1), v2)
    ElseIf Dimension(v1) = 1 And Dimension(v2) <> 1 Then
        catV = catV(v1, vector(v2))
    Else
        catV = catV(vector(v1), vector(v2))
    End If
End Function
    Function p_catV(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_catV = make_funPointer(AddressOf catV, firstParam, secondParam)
    End Function

'ベクトルを結合（可変長引数）
Function catVs(ParamArray vectors() As Variant) As Variant
    catVs = foldl1(p_catV, VBA.Array(vectors)(0))
End Function

'行方向に結合
Function catR(ByRef matrix1 As Variant, ByRef matrix2 As Variant) As Variant
    Dim i As Long, counter As Long
    Dim ret As Variant
    
    If Dimension(matrix1) <> 2 Or Dimension(matrix2) <> 2 Or colSize(matrix1) <> colSize(matrix2) Then
        catR = VBA.Array()
    Else
        ReDim ret(0 To rowSize(matrix1) + rowSize(matrix2) - 1, 0 To colSize(matrix1) - 1)
        counter = 0
        For i = LBound(matrix1, 1) To UBound(matrix1, 1) Step 1
            Call fillRow_imple(ret, counter, matrix1, i)
            counter = counter + 1
        Next i
        For i = LBound(matrix2, 1) To UBound(matrix2, 1) Step 1
            Call fillRow_imple(ret, counter, matrix2, i)
            counter = counter + 1
        Next i
        catR = ret
    End If
End Function

'列方向に結合
Function catC(ByRef matrix1 As Variant, ByRef matrix2 As Variant) As Variant
    Dim i As Long, counter As Long
    Dim ret As Variant
    
    If Dimension(matrix1) <> 2 Or Dimension(matrix2) <> 2 Or rowSize(matrix1) <> rowSize(matrix2) Then
        catC = VBA.Array()
    Else
        ReDim ret(0 To rowSize(matrix1) - 1, 0 To colSize(matrix1) + colSize(matrix2) - 1)
        counter = 0
        For i = LBound(matrix1, 2) To UBound(matrix1, 2) Step 1
            Call fillCol_imple(ret, counter, matrix1, i)
            counter = counter + 1
        Next i
        For i = LBound(matrix2, 2) To UBound(matrix2, 2) Step 1
            Call fillCol_imple(ret, counter, matrix2, i)
            counter = counter + 1
        Next i
        catC = ret
    End If
End Function

'配列の行/列の転置
Function transpose(ByRef matrix As Variant) As Variant
    Dim i As Long, j As Long, r As Long, c As Long
    Dim ret As Variant
    
    Select Case Dimension(matrix)
    Case 0
        transpose = matrix
    Case 1
        If LBound(matrix, 1) > UBound(matrix, 1) Then transpose = VBA.Array(): Exit Function
        transpose = matrix
    Case 2
        r = LBound(matrix, 1)
        c = LBound(matrix, 2)
        ReDim ret(0 To UBound(matrix, 2) - c, 0 To UBound(matrix, 1) - r)
        For i = 0 To UBound(matrix, 2) - c
            For j = 0 To UBound(matrix, 1) - r
                ret(i, j) = matrix(j + r, i + c)
            Next j
        Next i
        transpose = ret
    Case Else
        transpose = VBA.Array()
    End Select
End Function

'ふたつの配列の対応する要素どうしをcatVしてジャグ配列を作る
Public Function zip(ByRef a As Variant, ByRef b As Variant) As Variant
    zip = zipWith(AddressOf catV, a, b)
End Function
    Function p_zip(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_zip = make_funPointer(AddressOf zip, firstParam, secondParam)
    End Function

'zip（可変長引数）
Function zipVs(ParamArray vectors() As Variant) As Variant
    zipVs = foldl1(p_zip, VBA.Array(vectors)(0))
End Function

'zipされたジャグ配列をほどいて複数の1次元配列または一つの2次元配列に展開する
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
            ret(j) = z
        Next j
    Else
        ReDim ret(0 To sizeof(vec) - 1, 0 To colLen - 1)
        counter = LBound(vec)
        For i = LBound(ret, 1) To UBound(ret, 1) Step 1
            Call fillRow(ret, i, vec(counter))
            counter = counter + 1
        Next i
    End If
    unzip = ret
End Function

' Array(a, b)作成
Function makePair(ByRef a As Variant, ByRef b As Variant) As Variant
    makePair = VBA.Array(a, b)
End Function
    Public Function p_makePair(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_makePair = make_funPointer(AddressOf makePair, firstParam, secondParam)
    End Function

'ベクトルの直積に関数を適用した行列を作る
Public Function product_set(ByVal pCallback As Long, ByRef a As Variant, ByRef b As Variant) As Variant
    product_set = unzip(mapF(p_mapF(, b), mapF(p_bind1st(pCallback), a)), 2)
End Function
