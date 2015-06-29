Attribute VB_Name = "test_module"
'test_module
'Copyright (c) 2015 mmYYmmdd
Option Explicit

Declare PtrSafe Function GetTickCount Lib "Kernel32.dll" () As Long

'***********************************************************************
'   テスト関数
'   Sub vbaUnit         基本機能いろいろ
'   Sub sortTest        型がバラバラで配列も含む配列をソートするテスト
'   Sub treeTest        型がバラバラで配列も含む木構造のテスト
'   Sub sortTest2       数を並び替えて可能な最大数を返すテスト
'   Sub segmentsTest    引数の部分文字列のリストを取り出す
'   Sub segmentsTest2   ラムダ式使用バージョン
'***********************************************************************

'2点間の距離
Function distance(ByRef x As Variant, ByRef y As Variant) As Variant
    distance = foldr1(p_plus, mapF(p_mult, zipWith(p_minus, x, y))) ^ 0.5
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
Function fibonacci(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
    fibonacci = Array(a(LBound(a) + 1), a(LBound(a)) + a(LBound(a) + 1))
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
        p_innerProduct = make_funPointer(AddressOf innerProduct, firstParam, secondParam)
    End Function

'行列積
Function matrixMult(ByRef a As Variant, ByRef b As Variant) As Variant
    matrixMult = product_set(p_innerProduct, mapF(p_selectRow(a), a_rows(a)), mapF(p_selectCol(b), a_cols(b)))
End Function
    Function p_matrixMult(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_matrixMult = make_funPointer(AddressOf matrixMult, firstParam, secondParam)
    End Function

'素数判定(val：判定対象の自然数、pm : 既存の素数列, valはpmの最大数の2乗を超えないこと)
Function isPrime(ByRef val As Variant, ByRef pm As Variant) As Variant
    Dim z As Variant
    For Each z In pm
        If val < z * z Then Exit For
        If val Mod z = 0 Then
            isPrime = 0
            Exit Function
        End If
    Next z
    isPrime = 1
End Function
    Function p_isPrime(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_isPrime = make_funPointer(AddressOf isPrime, firstParam, secondParam)
    End Function

'ニュートン法による求根の１ステップ　：　x1 から x2 を出力する
'第1引数 ：　x ,  第2引数 (f, df/dx)
Function Newton_Raphson(ByRef x As Variant, ByRef fdf As Variant) As Variant
    Newton_Raphson = x - applyFun(x, fdf(0)) / applyFun(x, fdf(1))
End Function
    Function p_Newton_Raphson(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_Newton_Raphson = make_funPointer(AddressOf Newton_Raphson, firstParam, secondParam)
    End Function

'多項式の微分（係数のみ）
Function poly_deriv(ByRef coef As Variant) As Variant
    poly_deriv = headN(zipWith(p_mult, coef, iota(sizeof(coef) - 1, 0)), sizeof(coef) - 1)
End Function


'整数を各桁の数字の合計で比較する
Function compareSS(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim i As Long, aa As Long, bb As Long, aaa As Long, bbb As Long
    aa = Abs(CLng(a))
    bb = Abs(CLng(b))
    Do While 0 < aa
        aaa = aaa + (aa Mod 10)
        aa = aa \ 10
    Loop
    Do While 0 < bb
        bbb = bbb + (bb Mod 10)
        bb = bb \ 10
    Loop
    compareSS = IIf(aaa < bbb, 1, 0)
End Function
    Function p_compareSS(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_compareSS = make_funPointer(AddressOf compareSS, firstParam, secondParam)
    End Function


'テスト関数
Sub vbaUnit()
    Dim N As Long
    Dim Points As Variant, m As Variant, z As Variant, pred As Variant
    Dim N100 As Variant, m3 As Variant, m5 As Variant, m15 As Variant
    Dim init As Double, r As Double
    
    Debug.Print "------- mapF ----------"
    Debug.Print "mapF(p_log, Array(1, 2, 3, 4, 5, 6, 7))"
    printM mapF(p_log, Array(1, 2, 3, 4, 5, 6, 7))
    Debug.Print "mapF(p_minus(3), iota(1,15))"
    printM mapF(p_minus(3), iota(1, 15))
    Debug.Print "mapF(p_minus(, 3), iota(1, 15))"
    printM mapF(p_minus(, 3), iota(1, 15))

    Debug.Print "------- zipWith ----------"
    Debug.Print "zipWith(p_plus, Array(1, 2, 3, 4, 5), Array(10, 100, 1000, 100, 10))"
    printM zipWith(p_plus, Array(1, 2, 3, 4, 5), Array(10, 100, 1000, 100, 10))
    
    Debug.Print "------- foldl ----------"
    Debug.Print "foldl(p_minus, 0, iota(1, 100))  = (...(((0-1)-2)-3)-...-100"
    Debug.Print foldl(p_minus, 0, iota(1, 100))
    
    Debug.Print "------- foldr ----------"
    Debug.Print "foldr(p_minus, 0, iota(1, 100))  = 1-(2-(3-...(99-(100-0)))...)"
    Debug.Print foldr(p_minus, 0, iota(1, 100))
    
    Debug.Print "------- 円周率を確率的に求める（2通り） ------------"
    N = 9999
    Points = zip(mapF(p_rnd(, 1), repeat(0, N)), mapF(p_rnd(, 1), repeat(0, N)))
    printM Array("π≒", 4 * count_if(p_less(, 1#), mapF(p_distance(, Array(0, 0)), Points)) / N)
    printM Array("π≒", 4 * repeat_while(0, p_true, p_plus(p_less(p_distance(p_makePair(p_rnd(0, 1), p_rnd(0, 1)), Array(0, 0)), 1#)), N) / N)
    
    Debug.Print "------- ロジスティック漸化式 ------------"
    N = 10
    init = 0.1: r = 3.754
    printM scanl_Funs(init, repeat(p_Logistic(, r), N))
         'scanl(p_applyFun, init, repeat(p_Logistic(, r), N)) に相当
    printM scanr_Funs(init, repeat(p_Logistic(, r), N))
         'scanr(p_setParam, init, repeat(p_Logistic(, r), N)) に相当

    Debug.Print "------- フィボナッチ数列（5通り） ------------"
    N = 15
    printM unzip(scanl(p_applyFun, Array(0, 1), repeat(p_fibonacci, N)), 1)(0)
    printM unzip(scanl_Funs(Array(0, 1), repeat(p_fibonacci, N)), 1)(0)
    printM unzip(scanl(p_applyFun2by2, Array(0, 1), repeat(Array(p_secondArg, p_plus), N)), 1)(0)
    printM unzip(generate_while(Array(0, 1), p_true, p_makePair(p_getNth(1), p_plus(p_getNth(0), p_getNth(1))), N), 1)(0)
    printM unzip(generate_while(Array(0, 1), p_true, p_applyFun2by2(, Array(p_secondArg, p_plus)), N), 1)(0)
    
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

    Debug.Print "------- ソート ------------"
    m = mapF(p_getCLng(p_rnd(, 30)), repeat(10, 30))
    Debug.Print "ソート前"
    printM m
    Debug.Print "昇順ソート"
    printM subM(m, sortIndex(m))
    Debug.Print "各桁の数字の合計で比較するファンクタでソート"
    printM subM(m, sortIndex_pred(m, p_compareSS))
    
    Debug.Print "------- 行列積 ------------"
    printM matrixMult(makeM(4, 3, iota(1, 12)), makeM(3, 4, iota(1, 12)))

    Debug.Print "------- 素数列（[2,3,5]からの生成を3回適用） ------------"
    m = Array(2, 3, 5)
    z = iota(2, m(UBound(m)) ^ 2)
        m = filterR(z, mapF(p_isPrime(, m), z))
        printM catVs(headN(m, 5), Array("・・・"), tailN(m, 5))
    z = iota(2, m(UBound(m)) ^ 2)
        m = filterR(z, mapF(p_isPrime(, m), z))
        printM catVs(headN(m, 5), Array("・・・"), tailN(m, 5))
    z = iota(2, m(UBound(m)) ^ 2)
        m = filterR(z, mapF(p_isPrime(, m), z))
        printM catVs(headN(m, 5), Array("・・・"), tailN(m, 5))

    Debug.Print "------- 単純なNewton法による多項式の根（2通り） ------------"
    Debug.Print " f(x) = 2x^3 + x^2 - 5x + 0.5 の零点"
    z = Array(2, 1, -5, 0.5)
    m = VBA.Array(p_poly(, z), p_poly(, poly_deriv(z)))
    Debug.Print "x0 = 0  -> x = " & foldl_Funs(0, repeat(p_Newton_Raphson(, m), 7))
    Debug.Print "x0 = 1  -> x = " & repeat_while(1, p_less(0.0001, p_abs(m(0), 0)), p_Newton_Raphson(, m))
    Debug.Print "x0 = -7 -> x = " & repeat_while(-7, p_less(0.0001, p_abs(m(0), 0)), p_Newton_Raphson(, m))

    Debug.Print "------- 条件によるFind ------------"
    Debug.Print "乱数列 ( [0.0～100.0] * 10000個 ) から 29.9超 29.99未満のものを探す"
    Points = mapF(p_rnd(0), repeat(100, 10000))
    pred = p_mult(p_greater(, 29.9), p_less(, 29.99))
    m = find_pred(pred, Points)
    If (IsNull(m)) Then Debug.Print "なし" Else Debug.Print Points(m) & " (index=" & m & ")"
End Sub


'型がバラバラで配列も含む比較関数
Function comp000(ByRef a As Variant, ByRef b As Variant) As Variant
    ' 型が違う場合
    If VarType(a) <> VarType(b) Then
        comp000 = IIf(VarType(a) < VarType(b), 1, 0)
    Else
        ' 配列の場合
        If IsArray(a) Then
            ' 次元が異なる場合
            If Dimension(a) <> Dimension(b) Then
                comp000 = IIf(Dimension(a) < Dimension(b), 1, 0)
            Else
                comp000 = IIf(sizeof(a) < sizeof(b), 1, 0)
            End If
        ElseIf IsNull(a) Then
            comp000 = 0
        Else    ' 数値または文字列
            comp000 = IIf(a < b, 1, 0)
        End If
    End If
End Function
    Function p_comp000(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_comp000 = make_funPointer(AddressOf comp000, firstParam, secondParam)
    End Function

'型がバラバラで配列も含む配列をソートするテスト
Sub sortTest()
    Dim m As Variant, z As Variant:    ReDim m(0 To 9)
    Debug.Print "型がバラバラで配列も含む配列のソート"
    m(0) = 75676786
    m(1) = "abc"
    m(2) = iota(1, 8)
    m(3) = "鳥小屋"
    m(4) = 6
    m(5) = makeM(2, 2, iota(1, 4))
    m(6) = 300
    m(7) = iota(1, 15)
    m(8) = "犬小屋"
    m(9) = makeM(2, 3, iota(1, 6))
    Debug.Print vbLf & "===============ソート前==============="
    For Each z In m
        Debug.Print "-+-+-+-+-+-"
        printM z
    Next z
    Debug.Print vbLf & "===============ソート後==============="
    m = subM(m, sortIndex_pred(m, p_comp000))
    For Each z In m
        Debug.Print "-+-+-+-+-+-"
        printM z
    Next z
End Sub

'========木構造のテスト======================
Function makeNode(ByRef key As Variant, ByRef val As Variant, Optional ByRef comp As Variant) As Variant
    Dim ret As Variant:     ReDim ret(0 To 4)
    ret(0) = key
    ret(1) = val
    ret(2) = Empty  'Left
    ret(3) = Empty  'Right
    ret(4) = IIf(IsMissing(comp), Empty, comp)
    makeNode = moveVariant(ret)
End Function

Function makeNode0(ByRef key As Variant, ByRef val As Variant) As Variant
    makeNode0 = makeNode(key, val)
End Function
    Function p_makeNode0(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_makeNode0 = make_funPointer(AddressOf makeNode0, firstParam, secondParam)
    End Function

    Private Function less_with(ByRef a As Variant, ByRef b As Variant, ByRef comp As Variant) As Boolean
        If IsEmpty(comp) Then
            less_with = (a < b)
        Else
            less_with = 0 <> unbind_invoke(comp, a, b)
        End If
    End Function

Function insertNode(ByRef node As Variant, ByRef tree As Variant) As Variant
    If IsEmpty(tree) Then
        insertNode = makeNode(node(0), node(1), node(4))
    Else
        If less_with(node(0), tree(0), node(4)) Then
            swapVariant tree(2), insertNode(node, tree(2))
        ElseIf less_with(tree(0), node(0), node(4)) Then
            swapVariant tree(3), insertNode(node, tree(3))
        Else
            tree(1) = node(1)
        End If
        insertNode = moveVariant(tree)
    End If
End Function
    Function p_insertNode(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_insertNode = make_funPointer(AddressOf insertNode, firstParam, secondParam)
    End Function

Function getNode(ByRef key As Variant, ByRef tree As Variant) As Variant
    If IsEmpty(tree) Then
        getNode = Empty
    Else
        If less_with(key, tree(0), tree(4)) Then
            getNode = getNode(key, tree(2))
        ElseIf less_with(tree(0), key, tree(4)) Then
            getNode = getNode(key, tree(3))
        Else
            getNode = tree(1)
        End If
    End If
End Function
    Function p_getNode(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getNode = make_funPointer(AddressOf getNode, firstParam, secondParam)
    End Function

'型がバラバラで配列も含む木構造のテスト (速度的に実用性は無し)
Sub treeTest()
    Dim nodes As Variant, tree As Variant, N As Long, t As Long
    Dim Dic As Variant, i As Long
    Debug.Print "==== 型がバラバラで配列も含むキーによる木構造のテスト ===="
    '===============ノードの集合===============
    nodes = Array(makeNode(75676786, "A", p_comp000) _
                , makeNode("abc", "B", p_comp000) _
                , makeNode(iota(1, 8), "C", p_comp000) _
                , makeNode("鳥小屋", "D", p_comp000) _
                , makeNode(6, "E", p_comp000) _
                , makeNode(makeM(2, 2, iota(1, 4)), "F", p_comp000) _
                , makeNode(300, "G", p_comp000) _
                , makeNode(iota(1, 15), "H", p_comp000) _
                , makeNode("犬小屋", "I", p_comp000) _
                , makeNode(makeM(2, 3, iota(1, 6)), "J", p_comp000) _
              )
    '===============畳み込みによる木の構築===============
    tree = foldr(p_insertNode, Empty, nodes)
    '===============キーの選択===============
    Debug.Print """abc"" => ";
    printM getNode("abc", tree)
    Debug.Print """犬小屋"" => ";
    printM getNode("犬小屋", tree)
    Debug.Print "iota(1, 8) => ";
    printM getNode(iota(1, 8), tree)
    '==========================================================
    N = 10000
    Debug.Print "==== 0～" & N & " ランダム整数キー ===="
    nodes = zipWith(p_makeNode0, mapF(p_getCLng(p_rnd(0)), repeat(N, N)), iota(1, N))
    t = GetTickCount
    tree = foldr(p_insertNode, Empty, nodes)
    Debug.Print GetTickCount - t & "ms"
    printM mapF(p_getNode(, tree), iota(0, 10))
    
    Set Dic = CreateObject("Scripting.Dictionary")
    t = GetTickCount
    For i = UBound(nodes) To LBound(nodes) Step -1
        Dic.Item(nodes(i)(0)) = nodes(i)(1)
    Next i
    Debug.Print GetTickCount - t & "ms  " & Dic.Count & "_Items"
    For i = 0 To 10 Step 1
        Debug.Print Dic.Item(i);
    Next i
    Debug.Print ""
    Set Dic = Nothing
End Sub

'数を並び替えて可能な最大数を返すテスト
Sub sortTest2()
    Dim comp4 As Variant, arr As Variant, tmp As Variant, result As Variant
    
    Debug.Print "==== 数を並び替えて可能な最大数を返すテスト ===="
    Debug.Print "1～99 の整数乱数を20個作る"
    arr = mapF(p_getCLng(p_rnd(, 99)), repeat(0, 20))
    printM arr
    Debug.Print ""
    '--------------------
    comp4 = p_less(p_getCLng(p_plus(ph_1, ph_2)), p_getCLng(p_plus(ph_2, ph_1)))
    Debug.Print "比較関数はこれ↓  f(x, y) = CLng(x & y) < CLng(y & x)  に相当する"
    Debug.Print "  p_less(p_getCLng(p_plus(ph_1, ph_2)), p_getCLng(p_plus(ph_2, ph_1)))"
    Debug.Print ""
    '--------------------
    tmp = mapF(p_cStr, arr)    ' 文字列化
    Debug.Print "比較関数でソート（逆順）"
    result = subM(tmp, reverse(sortIndex_pred(tmp, comp4)))
    printM result
    Debug.Print "解"
    Debug.Print "  ";: Debug.Print foldl1(p_plus, result)     ' 文字列を結合して表示
End Sub

'====================================================================
' \a b -> [a] : map (a:) b のうち、map (a:) b の部分
Function consMap(ByRef a As Variant, ByRef v As Variant) As Variant
    consMap = mapF(p_cons(a), v)
End Function
    Public Function p_consMap(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_consMap = make_funPointer(AddressOf consMap, firstParam, secondParam)
    End Function

'引数の部分文字列のリストを取り出す「関数プログラミング実践入門」の問題
Sub segmentsTest()
    Debug.Print "==== 引数の部分文字列のリストを取り出すHaskell関数 ===="
    Debug.Print "    segments :: [a] -> [[a]]"
    Debug.Print "    segments = foldr (++) [] . scanr (\a b -> [a] : map (a:) b) []"
    Debug.Print "==== segmentsTest（ヘルパ関数使用バージョン） ===="
    Dim a As Variant, f As Variant, m As Variant
    
    a = Array("A", "B", "C", "D", "E")
    Debug.Print "これをfoldrする"
    Debug.Print "scanr(p_cons(p_makeSole, p_consMap(ph_1, ph_2))"
    f = p_cons(p_makeSole, p_consMap(ph_1, ph_2))
    m = foldr(p_catV, Array(), scanr(f, Array(), a))
    Debug.Print "Array(""A"", ""B"", ""C"", ""D"", ""E"") を展開する"
    printM mapF(p_join(, ""), m)
End Sub


'引数の部分文字列のリストを取り出す「関数プログラミング実践入門」の問題
'ラムダ式 lambdaExprによって、アドホックな関数 consMap を使用しなくて済むようにした
Sub segmentsTest2()
    Debug.Print "==== 引数の部分文字列のリストを取り出すHaskell関数 ===="
    Debug.Print "    segments :: [a] -> [[a]]"
    Debug.Print "    segments = foldr (++) [] . scanr (\a b -> [a] : map (a:) b) []"
    Debug.Print "==== segmentsTest2（ラムダ式バージョン） ===="
    Dim a As Variant, f As Variant, m As Variant
    
    a = Array("A", "B", "C", "D", "E")
    Debug.Print "これをfoldrする"
    Debug.Print "scanr(p_cons(p_makeSole, p_Mapf(lambdaExpr(p_cons, 1, ph_1), ph_2))"
    f = p_cons(p_makeSole, p_mapF(lambdaExpr(p_cons, 1, ph_1), ph_2))
    m = foldr(p_catV, Array(), scanr(f, Array(), a))
    Debug.Print "Array(""A"", ""B"", ""C"", ""D"", ""E"") を展開する"
    printM mapF(p_join(, ""), m)
End Sub

