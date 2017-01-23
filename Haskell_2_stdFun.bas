'Haskell_2_stdFun
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'********************************************************************
'   要素アクセス
' Sub       assignVar       汎用の変数コピー
' Function  firstArg        1番目の引数
' Function  secondArg       2番目の引数
' Function  getNth          N番目の配列要素取得（LBoundを無視した絶対位置）
' Function  getNth_b        N番目の配列要素取得（LBound基準）
' Sub       setNth_b        N番目の配列要素設定（LBound基準）
' Function  setNth_move     N番目の配列要素設定（LBoundを無視した絶対位置）
' Function  setNth_b_move   N番目の配列要素設定（LBound基準）
' Function  move_many       複数（可変長）の変数をmoveしてひとつのジャグ配列にする
' Sub       move_back       ジャグ配列から複数（可変長）の変数にmove back
'　-----------------------------------------------------------------
'     ファンクタ等　～
'********************************************************************

' 汎用の変数コピー
Public Sub assignVar(ByRef target As Variant, ByRef source As Variant)
    If IsObject(source) Then
        Set target = source
    Else
        target = source
    End If
End Sub


'1番目の引数
Function firstArg(ByRef a As Variant, ByRef b As Variant) As Variant
    Call assignVar(firstArg, a)
End Function
    Function p_firstArg(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_firstArg = make_funPointer(AddressOf firstArg, firstParam, secondParam)
    End Function

'2番目の引数
Function secondArg(ByRef a As Variant, ByRef b As Variant) As Variant
    Call assignVar(secondArg, b)
End Function
    Function p_secondArg(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_secondArg = make_funPointer(AddressOf secondArg, firstParam, secondParam)
    End Function

'N番目の配列要素取得（LBoundを無視した絶対位置）
Function getNth(ByRef vec As Variant, ByRef index As Variant) As Variant
    Call assignVar(getNth, vec(index))
End Function
    Function p_getNth(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getNth = make_funPointer(AddressOf getNth, firstParam, secondParam)
    End Function

'N番目の配列要素取得（LBound基準）
'index < 0 の場合は後ろから取得
Function getNth_b(ByRef vec As Variant, ByRef index As Variant) As Variant
    If 0 <= index Then
        Call assignVar(getNth_b, vec(index + LBound(vec)))
    Else
        Call assignVar(getNth_b, vec(UBound(vec) + 1 + index))
    End If
End Function
    Function p_getNth_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getNth_b = make_funPointer(AddressOf getNth_b, firstParam, secondParam)
    End Function

'N番目の配列要素設定（LBound基準）
'index < 0 の場合は後ろに設定
Sub setNth_b(ByRef vec As Variant, ByVal index As Long, ByRef value As Variant)
    If 0 <= index Then
        Call assignVar(vec(index + LBound(vec)), value)
    Else
        Call assignVar(vec(index + 1 + UBound(vec)), value)
    End If
End Sub

Function setNth_move(ByRef vec As Variant, ByVal index As Long, ByRef value As Variant)
    Call assignVar(vec(index), value)
    setNth_move = moveVariant(vec)
End Function

Function setNth_b_move(ByRef vec As Variant, ByVal index As Long, ByRef value As Variant)
    Call setNth_b(vec, index, value)
    setNth_b_move = moveVariant(vec)
End Function

' 複数の変数をmoveしてひとつのジャグ配列にする
Function move_many(ParamArray m() As Variant) As Variant
    If LBound(m) <= UBound(m) Then
        Dim ret As Variant
        ReDim ret(0 To UBound(m) - LBound(m))
        Dim i As Long, k As Long: k = 0
        For i = LBound(m) To UBound(m) Step 1
            swapVariant m(i), ret(k)
            k = k + 1
        Next i
    End If
    swapVariant move_many, ret
End Function

' ジャグ配列から複数（可変長）の変数にmove back
Sub move_back(ByRef m As Variant, ParamArray ret() As Variant)
    Dim i As Long, k As Long: k = LBound(ret)
    For i = LBound(m) To UBound(m) Step 1
        swapVariant m(i), ret(k)
        k = k + 1
    Next i
    m = Empty
End Sub

'********************************************************************
'     ファンクタ等
'   Function rowSize        配列の行数
'   Function colSize        配列の列数
'   Function sizeof         配列の全要素数または特定の軸の要素数
'   Function p_constant     定数関数
'   Function p_true         定数関数(true)
'   Function p_false        定数関数(false)
' * Function if_else        if else 選択
'   Function replaceNull    Nullを他の値に置換する
'   Function replaceEmpty   Emptyを他の値に置換する
'   Function expN           指数関数
'   Function logN           対数関数
'   Function absD           絶対値
'   Function plus           加算
'   Function minus          減算
'   Function mult           乗算
'   Function divide         除算
'   Function poly           多項式
'   Function min_fun        min
'   Function max_fun        max
'   Function CLng_          CLng（整数化）
'   Function CDbl_          CDbl（実数化）
'   Function CStr_          CStr（文字列化）
'   Function str_len        Len
'   Function str_left       Left（負の引数も可）
'   Function str_right      Right（負の引数も可）
'   Function str_mid        Mid
'   Function str_cat        文字列結合
'   Function splitFun       Split
'   Function joinFun        Join
'   Function gcm            gcm
'   Function lcm            lcm
'   Function equal          述語 Equal
'   Function notEqual       述語 Not Equal
'   Function less           述語 less
'   Function less_equal     述語 less_equal
'   Function greater        述語 greater
'   Function greater_equal  述語 greater_equal
'   Function is_null        述語 is_null
'   Function is_empty       述語 is_empty
'   Function is_valid       述語 is_valid
'********************************************************************

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

'配列の全要素数または特定の軸の要素数
Public Function sizeof(ByRef data As Variant, Optional ByVal axis As Long = 0) As Long
    Dim d As Long:  d = Dimension(data)
    Dim i As Long
    sizeof = 1
    If axis = 0 Then
        For i = 1 To d Step 1
            sizeof = sizeof * (1 + UBound(data, i) - LBound(data, i))
        Next i
    ElseIf 0 < axis And axis <= d Then
        sizeof = 1 + UBound(data, axis) - LBound(data, axis)
    End If
End Function
    
    Public Function p_sizeof(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_sizeof = make_funPointer_with_2nd_Default(AddressOf sizeof_, firstParam, secondParam)
    End Function
    Function sizeof_(ByRef data As Variant, Optional ByRef d As Variant) As Variant
        If IsNumeric(d) Then
            sizeof_ = sizeof(data, d)
        Else
            sizeof_ = sizeof(data)
        End If
    End Function

'定数関数
Function p_constant(ByRef x As Variant) As Variant
    p_constant = p_firstArg(x, 0)
End Function

'定数関数(true)
Function p_true() As Variant
    p_true = p_constant(1&)
End Function

'定数関数(false)
Function p_false() As Variant
    p_false = p_constant(0&)
End Function

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
            Call assignVar(if_else, trans(1 + lb))
        End If
    Else
        If is_bindFun(trans(2 + lb)) Then
            if_else = applyFun(val, trans(2 + lb))
        Else
            Call assignVar(if_else, trans(2 + lb))
        End If
    End If
    If is_placeholder(if_else) Then Call assignVar(if_else, val)
End Function
    Function p_if_else(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_if_else = make_funPointer(AddressOf if_else, firstParam, secondParam)
    End Function

'Nullを他の値に置換する
Function replaceNull(ByRef x As Variant, ByRef alt As Variant) As Variant
    If IsNull(x) Then
        Call assignVar(replaceNull, alt)
    Else
        Call assignVar(replaceNull, x)
    End If
End Function
    Function p_replaceNull(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_replaceNull = make_funPointer(AddressOf replaceNull, firstParam, secondParam)
    End Function

'Emptyを他の値に置換する
Function replaceEmpty(ByRef x As Variant, ByRef alt As Variant) As Variant
    If IsEmpty(x) Then
        Call assignVar(replaceEmpty, alt)
    Else
        Call assignVar(replaceEmpty, x)
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
Function logN(ByRef a As Variant, Optional ByRef base As Variant) As Variant
    If IsMissing(base) Then
        logN = Log(a)
    Else
        logN = Log(a) / Log(base)
    End If
End Function
    Function p_log(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_log = make_funPointer_with_2nd_Default(AddressOf logN, firstParam, secondParam)
    End Function

'絶対値
Function absD(ByRef val As Variant, Optional ByRef dummy As Variant) As Variant
    If IsMissing(dummy) Then dummy = 0
    absD = Abs(val - dummy)
End Function
    Function p_abs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_abs = make_funPointer_with_2nd_Default(AddressOf absD, firstParam, secondParam)
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
    poly = 0#
    Dim i As Long
    For i = LBound(coef) To UBound(coef) Step 1
        poly = poly * x + coef(i)
    Next i
End Function
    Function p_poly(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_poly = make_funPointer(AddressOf poly, firstParam, secondParam)
    End Function

'min
Function min_fun(ByRef a As Variant, ByRef b As Variant) As Variant
    min_fun = IIf(a < b, a, b)
End Function
    Function p_min(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_min = make_funPointer(AddressOf min_fun, firstParam, secondParam)
    End Function

'max
Function max_fun(ByRef a As Variant, ByRef b As Variant) As Variant
    max_fun = IIf(a < b, b, a)
End Function
    Function p_max(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_max = make_funPointer(AddressOf max_fun, firstParam, secondParam)
    End Function
    
'CLng
Function CLng_(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    CLng_ = 0
    If IsNumeric(a) Then CLng_ = CLng(a)
    If IsDate(a) Then CLng_ = CLng(DateValue(a))
End Function
    Function p_CLng(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_CLng = make_funPointer(AddressOf CLng_, firstParam, secondParam)
    End Function

'CDbl
Function CDbl_(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    CDbl_ = 0#
    If IsNumeric(a) Then CDbl_ = CDbl(a)
    If IsDate(a) Then CDbl_ = CDbl(DateValue(a))
End Function
    Function p_CDbl(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_CDbl = make_funPointer(AddressOf CDbl_, firstParam, secondParam)
    End Function

'CStr
Function CStr_(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    CStr_ = ""
    If Not IsArray(a) And Not IsObject(a) And Not IsNull(a) Then CStr_ = CStr(a)
End Function
    Function p_CStr(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_CStr = make_funPointer(AddressOf CStr_, firstParam, secondParam)
    End Function

'Len
Function str_len(ByRef st As Variant, Optional ByRef dummy As Variant) As Variant
    str_len = 0
    If VarType(st) = vbString Then str_len = Len(st)
End Function
    Function p_len(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_len = make_funPointer(AddressOf str_len, firstParam, secondParam)
    End Function
    
'Left
Function str_left(ByRef st As Variant, ByRef length As Variant) As Variant
    str_left = left(st, length)
End Function
    Function p_left(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_left = make_funPointer(AddressOf str_left, firstParam, secondParam)
    End Function
    
'Right
Function str_right(ByRef st As Variant, ByRef length As Variant) As Variant
    str_right = right(st, length)
End Function
    Function p_right(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_right = make_funPointer(AddressOf str_right, firstParam, secondParam)
    End Function
    
'Mid
Function str_mid(ByRef st As Variant, ByRef begin_len As Variant) As Variant
    str_mid = mid(st, begin_len(0), begin_len(1))
End Function
    Function p_mid(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_mid = make_funPointer(AddressOf str_mid, firstParam, secondParam)
    End Function

'文字列結合
Function str_cat(ByRef s1 As Variant, ByRef s2 As Variant) As Variant
    str_cat = CStr_(s1) & CStr_(s2)
End Function
    Function p_str_cat(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_str_cat = make_funPointer(AddressOf str_cat, firstParam, secondParam)
    End Function

'Split
Function splitFun(ByRef s As Variant, ByRef delim As Variant) As Variant
    splitFun = Split(s, delim)
End Function
    Function p_split(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_split = make_funPointer(AddressOf splitFun, firstParam, secondParam)
    End Function

'Join
Function joinFun(ByRef m As Variant, ByRef delim As Variant) As Variant
    If IsEmpty(m) Or IsNull(m) Then
        joinFun = ""
    Else
        joinFun = join(m, delim)
    End If
End Function
    Function p_join(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_join = make_funPointer(AddressOf joinFun, firstParam, secondParam)
    End Function

'gcm
Function gcm(ByRef a As Variant, ByRef b As Variant) As Variant
    If a = 0 Then
        gcm = b
    ElseIf b = 0 Then
        gcm = a
    Else
        gcm = gcm(b, a Mod b)
    End If
End Function
    Function p_gcm(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_gcm = make_funPointer(AddressOf gcm, firstParam, secondParam)
    End Function
    
'lcm
Function lcm(ByRef a As Variant, ByRef b As Variant) As Variant
    lcm = a * b / gcm(a, b)
End Function
    Function p_lcm(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_lcm = make_funPointer(AddressOf lcm, firstParam, secondParam)
    End Function
    
'述語 equal
Function equal(ByRef a As Variant, ByRef b As Variant) As Variant
    If IsNull(a) Or IsNull(b) Then
        equal = IIf(IsNull(a) = IsNull(b), 1, 0)
    Else
        equal = IIf(a = b, 1, 0)
    End If
End Function
    Function p_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal = make_funPointer(AddressOf equal, firstParam, secondParam)
    End Function

'述語 not equal
Function notEqual(ByRef a As Variant, ByRef b As Variant) As Variant
    notEqual = IIf(equal(a, b), 0, 1)
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

'述語 is_null
Function is_null(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    is_null = IIf(IsNull(a), 1&, 0&)
End Function
    Function p_is_null(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_is_null = make_funPointer(AddressOf is_null, firstParam, secondParam)
    End Function

'述語 is_empty
Function is_empty(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    is_empty = IIf(IsEmpty(a), 1&, 0&)
End Function
    Function p_is_empty(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_is_empty = make_funPointer(AddressOf is_empty, firstParam, secondParam)
    End Function

'述語 is_valid
Function is_valid(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    is_valid = IIf(IsEmpty(a) Or IsNull(a), 0&, 1&)
End Function
    Function p_is_valid(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_is_valid = make_funPointer(AddressOf is_valid, firstParam, secondParam)
    End Function
