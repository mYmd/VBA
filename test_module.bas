Attribute VB_Name = "test_module"
'test_module
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'2点間の距離
Function distance(ByRef x As Variant, Optional ByRef y As Variant) As Variant
    distance = foldr1(p_plus, mapF(p_poly(, Array(1, 0, 0)), zipWith(p_minus, x, IIf(IsMissing(y), repeat(0, sizeof(x)), y)))) ^ 0.5
End Function
    Function p_distance(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_distance = make_funPointer(AddressOf distance, firstParam, secondParam)
    End Function

'乱数
Function vbRand(ByRef from_ As Variant, ByRef to_ As Variant) As Variant
    vbRand = (to_ - from_) * Rnd() + from_
End Function
    Function p_rnd(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_rnd = make_funPointer(AddressOf vbRand, firstParam, secondParam)
    End Function

'=======================================================================
'フィボナッチ関数   (0,1)->(1,1)->(1,2)->(2,3)->(3,5)->(5,8)-> ...
Function fibonacci(ByRef a As Variant, ByRef b As Variant) As Variant
    fibonacci = Array(b, a + b)
End Function
    Function p_fibonacci(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_fibonacci = make_funPointer(AddressOf fibonacci, firstParam, secondParam)
    End Function

'ロジスティック写像   (Xn, r)->Xn+1
Function Logistic(ByRef x As Variant, ByRef r As Variant) As Variant
    Logistic = r * x * (1 - x)
End Function
    Function p_Logistic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_Logistic = make_funPointer(AddressOf Logistic, firstParam, secondParam)
    End Function

'円／球
Function circle_(ByRef point As Variant, ByRef r As Variant) As Variant
    circle_ = VBA.Array(point, r)
End Function
    Function p_circle(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_circle = make_funPointer(AddressOf circle_, firstParam, secondParam)
    End Function

'円の面積／球の体積
Function circleArea(ByRef circle__ As Variant, ByRef dummy As Variant) As Variant
    Select Case sizeof(circle__(0))
    Case 1
        circleArea = 2 * Abs(circle__(1))
    Case 2
        circleArea = 4 * Atn(1) * circle__(1) ^ 2
    Case 3
        circleArea = (4 / 3#) * 4 * Atn(1) * circle__(1) ^ 3
    Case Else
        circleArea = 0
    End Select
End Function
    Function p_circleArea(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_circleArea = make_funPointer(AddressOf circleArea, firstParam, secondParam)
    End Function

'内積inner product
Function innerProduct(ByRef a As Variant, ByRef b As Variant) As Variant
    innerProduct = foldl1(p_plus, zipWith(p_mult, a, b))
End Function
    Function p_innerProduct(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_innerPproduct = make_funPointer(AddressOf innerProduct, firstParam, secondParam)
    End Function

'行列積
Function matrixMult(ByRef a As Variant, ByRef b As Variant) As Variant
    matrixMult = product_set(AddressOf innerProduct, mapF(p_selectRow(a), a_rows(a)), mapF(p_selectCol(b), a_cols(b)))
End Function
    Function p_matrixMult(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_matrixMult = make_funPointer(AddressOf matrixMult, firstParam, secondParam)
    End Function


Sub vbaUnit()
    Dim N As Long
    Dim Points As Variant, m As Variant, z As Variant
    Dim N100 As Variant, m3 As Variant, m5 As Variant, m15 As Variant
    Dim init As Double, r As Double
    
    Debug.Print "------- mapF ----------"
    Debug.Print "mapF(p_log, Array(1, 2, 3, 4, 5, 6, 7))"
    printM mapF(p_log, Array(1, 2, 3, 4, 5, 6, 7))
    Debug.Print "mapF(p_minus(3), iota(1,15))"
    printM mapF(p_minus(3), iota(1, 15))
    Debug.Print "mapF(p_minus(, 3), iota(1, 15))"
    printM mapF(p_minus(, 3), iota(1, 15))

    Debug.Print "------- mapFのネスト ----------"
    m = mapF(p_mapF(, iota(1, 10)), Array(p_log, p_plus(100), p_divide(, 100)))
    Debug.Print "m = mapF(p_mapF(, iota(1, 10)), Array(p_log, p_plus(100), p_divide(, 100)))"
    Debug.Print "m(0)"
    printM m(0)
    Debug.Print "m(1)"
    printM m(1)
    Debug.Print "m(2)"
    printM m(2)
    
    Debug.Print "------- zipWith ----------"
    Debug.Print "zipWith(p_plus, Array(1, 2, 3, 4, 5), Array(10, 100, 1000, 100, 10))"
    printM zipWith(p_plus, Array(1, 2, 3, 4, 5), Array(10, 100, 1000, 100, 10))
    
    Debug.Print "------- foldl ----------"
    Debug.Print "foldl(p_minus, 0, iota(1, 100))  = (...(((0-1)-2)-3)-...-100"
    Debug.Print foldl(p_minus, 0, iota(1, 100))
    
    Debug.Print "------- foldr ----------"
    Debug.Print "foldr(p_minus, 0, iota(1, 100))  = 1-(2-(3-...(99-(100-0)))...)"
    Debug.Print foldr(p_minus, 0, iota(1, 100))
    
    Debug.Print "------- 円周率を確率的に求める ------------"
    N = 10000
    Points = zip(mapF(p_rnd(, 1), repeat(0, N)), mapF(p_rnd(, 1), repeat(0, N)))
    printM Array("π≒", 4 * count_if(p_less(, 1#), mapF(p_distance, Points)) / N)
    
    Debug.Print "------- ロジスティック漸化式 ------------"
    N = 10
    init = 0.1: r = 3.754
    printM scanl_Funs(init, repeat(p_Logistic(, r), N))
         'scanl(p_applyFun, init, repeat(p_Logistic(, r), N)) に相当
    printM scanr_Funs(init, repeat(p_Logistic(, r), N))
         'scanr(p_setParam, init, repeat(p_Logistic(, r), N)) に相当

    Debug.Print "------- フィボナッチ数列（４通り） ------------"
    'bindFunを使っている理由は、applyFunの適用例 7. もしくは setParamの適用例 4. を参照
    N = 10
    printM unzip(scanl(p_applyFun, Array(0, 1), repeat(bindFun(p_fibonacci), N)), 1)(0)
    printM unzip(scanl_Funs(Array(0, 1), repeat(bindFun(p_fibonacci), N)), 1)(0)
    printM unzip(scanr_Funs(Array(0, 1), repeat(bindFun(p_fibonacci), N)), 1)(0)
    printM unzip(scanl(p_applyFun2by2, Array(0, 1), repeat(Array(p_secondArg, p_plus), N)), 1)(0)

    Debug.Print "------- FizzBuzz ------------"
    m = Array(Array(p_mod(, 15), Null, "FizzBuzz"), _
              Array(p_mod(, 5), Null, "Buzz"), _
              Array(p_mod(, 3), placeholder, "Fizz"))
    printM foldl1(p_replaceNull, product_set(p_if_else, iota(1, 100), m), 2)

    Debug.Print "------- zip ------------"
    m = "文字をひとつずつ分離する"
    printM mapF(p_mid(m), zip(iota(1, Len(m)), repeat(1, Len(m))))
    
    m = zip(Array(1, 2, 3, 4, 5), Array(11, 12, 13, 14, 15))
    For Each z In m: printM z: Next z

    Debug.Print "------- unzip ------------"
    printM unzip(m, 1)(0)
    printM unzip(m, 1)(1)
    printM unzip(m, 2)

    Debug.Print "------- ソートとequal_range ------------"
    m = mapF(p_getCLng, mapF(p_rnd(, 10), repeat(0, 20)))
    Debug.Print "ソート前"
    printM m
    m = subM(m, sortIndex(m))
    Debug.Print "ソート後"
    printM m
    z = equal_range(m, 5)
    Debug.Print "equal_range(m, 5)"
    printM mapF(p_getNth(, m), iota(z(0), z(1)))
    
    Debug.Print "------- 行列積 ------------"
    printM matrixMult(makeM(4, 3, iota(1, 12)), makeM(3, 4, iota(1, 12)))
End Sub
