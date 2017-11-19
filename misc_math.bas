Attribute VB_Name = "misc_math"
'misc_math
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'********************************************************************
'   数学関数 (Haskell_2_stdFunにも一部の数学的関数がある)

' Function  isPrimeNumber       素数判定
' Function  primeNumbers        素数一覧
' Function  sin_fun(p_sin)      Sin
' Function  cos_fun(p_con)      Cos
' Function  pow_fun(p_pow)      Pow
' Function  make_polyCoef       多項式の微分または不定積分（係数の生成）
' Function  newton_method       ニュートン法による求根（の1ステップ）
' Function  integral_simpson    シンプソン法による数値積分
' Function  make_complex        複素数の生成
'********************************************************************

    Private primeNumbers_() As Long
    Private syntheticFlag_() As Boolean
    Private max_nval_ As Long

' 素数判定
Public Function isPrimeNumber(ByVal val As Long) As Boolean
    If max_nval_ = 0 Then
        ReDim primeNumbers_(0 To 1)
        primeNumbers_(0) = 2: primeNumbers_(1) = 3
        ReDim syntheticFlag_(0 To 3)
        syntheticFlag_(0) = True: syntheticFlag_(1) = True
        max_nval_ = 3
    End If
    Do While max_nval_ < val
        enlargePrime val
        max_nval_ = UBound(syntheticFlag_)
    Loop
    If val < 0 Then
        isPrimeNumber = False
    Else
        isPrimeNumber = Not syntheticFlag_(val)
    End If
End Function

' 素数一覧
Public Function primeNumbers(Optional ByVal val As Long = -1) As Variant
    Call isPrimeNumber(val)
    primeNumbers = primeNumbers_
End Function

    Private Sub enlargePrime(ByVal val As Long)
        Dim lastPrime As Long
        lastPrime = primeNumbers_(UBound(primeNumbers_))
        Dim flag_end As Long, flag_end_ex As Long
        flag_end = UBound(syntheticFlag_)
        flag_end_ex = min_fun(val, lastPrime ^ 2)
        ReDim Preserve syntheticFlag_(0 To flag_end_ex)
        Dim counter As Long:    counter = flag_end_ex - flag_end
        Dim p_iter As Long, i As Long
        For p_iter = 0 To UBound(primeNumbers_) Step 1
            Dim prime_ As Long:     prime_ = primeNumbers_(p_iter)
            For i = prime_ * (1 + flag_end \ prime_) To flag_end_ex Step prime_
                If Not syntheticFlag_(i) Then counter = counter - 1
                syntheticFlag_(i) = True
            Next i
        Next p_iter
        p_iter = UBound(primeNumbers_)   '
        ReDim Preserve primeNumbers_(0 To p_iter + counter)
        For i = flag_end + 1 To flag_end_ex Step 1
            If syntheticFlag_(i) = 0 Then
                p_iter = p_iter + 1
                primeNumbers_(p_iter) = i
            End If
        Next i
    End Sub

' Sin
Function sin_fun(ByRef x As Variant, ByRef dummy As Variant) As Variant
    sin_fun = Sin(x)
End Function
    Function p_sin(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_sin = make_funPointer(AddressOf sin_fun, firstParam, secondParam)
    End Function

' Cos
Function cos_fun(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
    cos_fun = Cos(x)
End Function
    Function p_cos(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_cos = make_funPointer(AddressOf cos_fun, firstParam, secondParam)
    End Function

' Pow
Function pow_fun(ByRef x As Variant, ByRef y As Variant) As Variant
    pow_fun = x ^ y
End Function
    Function p_pow(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_pow = make_funPointer(AddressOf pow_fun, firstParam, secondParam)
    End Function

' 多項式の微分または不定積分（係数の生成）
' 多項式そのものは Haskell_2_stdFun::poly
Function make_polyCoef(ByRef coef As Variant, Optional ByRef deriv_N As Variant) As Variant
    Dim i As Long, dimen As Long, derivN As Long
    dimen = sizeof(coef) - 1
    derivN = IIf(IsMissing(deriv_N), 0, min_fun(dimen + 1, CLng(deriv_N)))
    Dim coefMatrix As Variant
    If 0 <= derivN Then     ' 微分
        coefMatrix = makeM(derivN + 1, dimen + 1, 0)
        Call fillRow(coefMatrix, 0, coef)
        For i = 1 To derivN Step 1
            Call fillRow(coefMatrix, i, iota(dimen - i + 1, 0))
        Next i
        make_polyCoef = headN(foldl1(p_mult, coefMatrix, 1), max_fun(1, dimen + 1 - derivN))
    Else    ' 不定積分
        coefMatrix = makeM(1 - derivN, dimen + 1, 0)
        Call fillRow(coefMatrix, 0, coef)
        For i = 1 To 0 - derivN Step 1
            Call fillRow(coefMatrix, i, iota(dimen + i, i))
        Next i
        make_polyCoef = catV(foldl1(p_divide, coefMatrix, 1), repeat(0, -derivN))
    End If
End Function
    Function p_make_polyCoef(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_make_polyCoef = make_funPointer(AddressOf make_polyCoef, firstParam, secondParam)
    End Function

'ニュートン法による求根（の1ステップ）　：　x1 から x2 を出力する
'第1引数 ：　x ,  第2引数 (f, df/dx)
Function newton_method(ByRef x As Variant, ByRef f_df As Variant) As Variant
    newton_method = x - applyFun(x, f_df(LBound(f_df))) / applyFun(x, f_df(UBound(f_df)))
End Function
    Function p_newton_method(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_newton_method = make_funPointer(AddressOf newton_method, firstParam, secondParam)
    End Function

' シンプソン法による数値積分
Function integral_simpson(ByRef fun As Variant, _
                          ByVal begin_ As Double, _
                          ByVal end_ As Double, _
                          ByVal N As Long) As Double
    Dim xs As Variant, ys As Variant
    xs = mapF(p_poly(, Array((end_ - begin_) / 2 / N, begin_)), iota(0, 2 * N))
    ys = mapF(fun, xs)
    Dim constants As Variant
    constants = makeM(2 * N + 1)
    Call fillPattern(constants, Array(2, 4))
    constants(0) = 1
    constants(2 * N) = 1
    integral_simpson = foldl1(p_plus, zipWith(p_mult, constants, ys)) * (end_ - begin_) / N / 6
End Function


' 複素数の生成
Function make_complex(ByRef r As Variant, ByRef i As Variant) As Variant
    make_complex = VBA.Array(CDbl(r), CDbl(i))
End Function
    Function p_make_complex(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_make_complex = make_funPointer(AddressOf make_complex, firstParam, secondParam)
    End Function

Function make_complex_polar(ByRef r As Variant, ByRef arg As Variant) As Variant
    make_complex_polar = VBA.Array(CDbl(r) * Cos(arg), CDbl(r) * Sin(arg))
End Function
    Function p_make_complex_polar(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_make_complex_polar = make_funPointer(AddressOf make_complex_polar, firstParam, secondParam)
    End Function
    
Function show_complex(ByRef C As Variant, Optional ByRef dummy As Variant) As Variant
    If C(1) < 0# Then
        show_complex = C(0) & C(1) & "i"
    Else
        show_complex = C(0) & "+" & C(1) & "i"
    End If
End Function
    Function p_show_complex(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_show_complex = make_funPointer(AddressOf show_complex, firstParam, secondParam)
    End Function

Function show_complex_polar(ByRef C As Variant, Optional ByRef dummy As Variant) As Variant
    show_complex_polar = "(" & complex_abs(C) & ", " & complex_arg(C) & ")"
End Function
    Function p_show_complex_polar(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_show_complex_polar = make_funPointer(AddressOf show_complex_polar, firstParam, secondParam)
    End Function

Function complex_add(ByRef a As Variant, ByRef b As Variant) As Variant
    complex_add = VBA.Array(a(0) + b(0), a(1) + b(1))
End Function
    Function p_complex_add(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_complex_add = make_funPointer(AddressOf complex_add, firstParam, secondParam)
    End Function

Function complex_minus(ByRef a As Variant, ByRef b As Variant) As Variant
    complex_minus = VBA.Array(a(0) - b(0), a(1) - b(1))
End Function
    Function p_complex_minus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_complex_minus = make_funPointer(AddressOf complex_minus, firstParam, secondParam)
    End Function

Function complex_mult(ByRef a As Variant, ByRef b As Variant) As Variant
    complex_mult = VBA.Array(a(0) * b(0) - a(1) * b(1), a(0) * b(1) + a(1) * b(0))
End Function
    Function p_complex_mult(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_complex_mult = make_funPointer(AddressOf complex_mult, firstParam, secondParam)
    End Function

Function complex_divide(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim d As Double
    d = b(0) ^ 2 + b(1) ^ 2
    complex_divide = VBA.Array((a(0) * b(0) + a(1) * b(1)) / d, (-a(0) * b(1) + a(1) * b(0)) / d)
End Function
    Function p_complex_divide(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_complex_divide = make_funPointer(AddressOf complex_divide, firstParam, secondParam)
    End Function

Function complex_cnj(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    complex_cnj = VBA.Array(a(0), -a(1))
End Function
    Function p_complex_cnj(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_complex_cnj = make_funPointer(AddressOf complex_cnj, firstParam, secondParam)
    End Function

Function complex_abs2(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    complex_abs2 = a(0) ^ 2 + a(1) ^ 2
End Function
    Function p_complex_abs2(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_complex_abs2 = make_funPointer(AddressOf complex_abs2, firstParam, secondParam)
    End Function

Function complex_abs(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    complex_abs = complex_abs2(a) ^ 0.5
End Function
    Function p_complex_abs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_complex_abs = make_funPointer(AddressOf complex_abs, firstParam, secondParam)
    End Function

Function complex_arg(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    complex_arg = IIf(a(0) = 0#, 0#, Atn(a(1) / a(0)))
    If a(0) < 0# Then
        complex_arg = complex_arg + 4 * Atn(1)
    ElseIf a(1) < 0# Then
        complex_arg = complex_arg + 8 * Atn(1)
    End If
End Function
    Function p_complex_arg(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_complex_arg = make_funPointer(AddressOf complex_arg, firstParam, secondParam)
    End Function
