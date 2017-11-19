Attribute VB_Name = "test_module"
'test_module
'Copyright (c) 2015 mmYYmmdd
Option Explicit

Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

'***********************************************************************
'   �e�X�g�֐�
'   Sub vbaUnit                 ��{�@�\���낢��(1)
'   Sub vbaUnit2                ��{�@�\���낢��(2)
'***********************************************************************

Public Sub vbaUnit()
    Call mapF_test:         Debug.Print vbCrLf
    Call zipWith_test:      Debug.Print vbCrLf
    Call foldl_foldr_test:  Debug.Print vbCrLf
    Call pi_test:           Debug.Print vbCrLf
    Call logistic_test:     Debug.Print vbCrLf
    Call fibonacci_test:    Debug.Print vbCrLf
    Call FizzBuzz_test:     Debug.Print vbCrLf
    Call zip_unzip_test:    Debug.Print vbCrLf
    Call sort_test:         Debug.Print vbCrLf
    Call matrixMult_test:   Debug.Print vbCrLf
    Call primeNumber_test:  Debug.Print vbCrLf
    Call newton_test:       Debug.Print vbCrLf
    Call find_test
End Sub

Public Sub vbaUnit2()
    Call flatten_test:  Debug.Print vbCrLf
    Call sort_test2:        Debug.Print vbCrLf
    Call sort_test3:            Debug.Print vbCrLf
    Call segmentsTest1:             Debug.Print vbCrLf
    Call segmentsTest2:             Debug.Print vbCrLf
    Call curiouslyRecursiveTest:    Debug.Print vbCrLf
    Call equal_range_test
End Sub

    ' ========== vbaUnit �̃e�X�g=================
    Private Sub mapF_test()
        Debug.Print "------- mapF ----------"
        Debug.Print "mapF(p_log, Array(1, 2, 3, 4, 5, 6, 7))"
        printM mapF(p_log, Array(1, 2, 3, 4, 5, 6, 7))
        Debug.Print "mapF(p_minus(3), iota(1,15))"
        printM mapF(p_minus(3), iota(1, 15))
        Debug.Print "mapF(p_minus(, 3), iota(1, 15))"
        printM mapF(p_minus(, 3), iota(1, 15))
    End Sub

    Private Sub zipWith_test()
        Debug.Print "------- zipWith ----------"
        Debug.Print "zipWith(p_plus, Array(1, 2, 3, 4, 5), Array(10, 100, 1000, 100, 10))"
        printM zipWith(p_plus, Array(1, 2, 3, 4, 5), Array(10, 100, 1000, 100, 10))
    End Sub

    Private Sub foldl_foldr_test()
        Debug.Print "------- foldl ----------"
        Debug.Print "foldl(p_minus, 0, iota(1, 100))  = (...(((0-1)-2)-3)-...-100"
        Debug.Print foldl(p_minus, 0, iota(1, 100))
        Debug.Print "------- foldr ----------"
        Debug.Print "foldr(p_minus, 0, iota(1, 100))  = 1-(2-(3-...(99-(100-0)))...)"
        Debug.Print foldr(p_minus, 0, iota(1, 100))
    End Sub
    
    Private Sub pi_test()
        Dim n As Long
        Dim points As Variant
        Debug.Print "------- �~�������m���I�ɋ��߂�i2�ʂ�j ------------"
        n = 9999
        points = zip(uniform_real_dist(n, 0, 1), uniform_real_dist(n, 0, 1))
        printM Array("�΁�", 4 * count_if(p_less(, 1#), mapF(p_distance2(, Array(0, 0)), points)) / n)
        printM Array("�΁�", 4 * repeat_while(0, p_true, p_plus(p_less(p_distance2(p_makePair(p_rnd(0, 1), p_rnd(0, 1)), Array(0, 0)), 1#)), n) / n)
    End Sub
    
    Private Sub logistic_test()
        Dim n As Long
        Dim init As Double, r As Double
        Debug.Print "------- ���W�X�e�B�b�N�Q���� ------------"
        n = 10
        init = 0.1: r = 3.754
        printM scanl_Funs(init, repeat(p_Logistic(, r), n))
             'scanl(p_applyFun, init, repeat(p_Logistic(, r), N)) �ɑ���
        printM scanr_Funs(init, repeat(p_Logistic(, r), n))
             'scanr(p_setParam, init, repeat(p_Logistic(, r), N)) �ɑ���
    End Sub

    Private Sub fibonacci_test()
        Dim n As Long: n = 15
        Dim f As Variant
        Debug.Print "------- �t�B�{�i�b�`����i2�ʂ�j ------------"
        printM unzip(scanl(p_applyFun, Array(0, 1), repeat(p_fibonacci, n)), 1)(0)
        f = p_push_back_move(, p_foldl1(p_plus, p_tailN(, 2)))
        printM repeat_while(Array(0, 1), p_true, f, n - 1)
        Debug.Print "------- �e�g���i�b�`���� ------------"
        f = p_push_back_move(, p_foldl1(p_plus, p_tailN(, 4)))
        printM repeat_while(Array(0, 0, 0, 1), p_true, f, n - 1)
    End Sub
    
    Private Sub FizzBuzz_test()
        Dim m As Variant
        Debug.Print "------- FizzBuzz�i3�ʂ�j ------------"
        m = Array(Array(p_mod(, 15), Null, "FizzBuzz"), _
                  Array(p_mod(, 5), Null, "Buzz"), _
                  Array(p_mod(, 3), placeholder, "Fizz"))
        printM foldl1(p_replaceNull, product_set(p_if_else, iota(1, 100), m), 2)
        
        m = p_if_else(, Array(p_mod(, 3), placeholder, "Fizz"))
        m = p_if_else(, Array(p_mod(, 5), m, "Buzz"))
        m = p_if_else(, Array(p_mod(, 15), m, "FizzBuzz"))
        printM mapF(m, iota(1, 100))
        
        printM mapF(p_try(p_mod(, 15), p_try(p_mod(, 5), p_try(p_mod(, 3), , "Fizz"), "Buzz"), "FizzBuzz"), _
                    iota(1, 100))
    End Sub

    Private Sub zip_unzip_test()
        Dim m As Variant, z As Variant
        Debug.Print "------- zip ------------"
        m = "�������ЂƂ���������"
        printM mapF(p_mid(m), zip(iota(1, Len(m)), repeat(1, Len(m))))
        m = zip(Array(1, 2, 3, 4, 5), Array(11, 12, 13, 14, 15))
        For Each z In m: printM z: Next z
        Debug.Print "------- unzip ------------"
        printM unzip(m, 1)(0)
        printM unzip(m, 1)(1)
        printM unzip(m, 2)
    End Sub

    Private Sub sort_test()
        Dim m As Variant
        Debug.Print "------- �\�[�g ------------"
        m = uniform_int_dist(30, 10, 30)
        Debug.Print "�\�[�g�O"
        printM m
        Debug.Print "�����\�[�g"
        printM subM(m, sortIndex(m))
        Debug.Print "�e���̐����̍��v�Ŕ�r����t�@���N�^�Ń\�[�g"
        printM subM(m, sortIndex_pred(m, p_compareSS))
    End Sub

    Private Sub matrixMult_test()
        Debug.Print "------- �s��� ------------"
        printM matrixMult_(makeM(4, 3, iota(1, 12)), makeM(3, 4, iota(1, 12)))
    End Sub
    
    Private Sub primeNumber_test()
        Dim m As Variant, z As Variant
        Debug.Print "------- �f����i[2,3,5]����̐�����3��K�p�j ------------"
        m = Array(2, 3, 5)
        z = iota(2, m(UBound(m)) ^ 2)
            m = filterR(z, mapF(p_isPrime_(, m), z))
            printM catVs(headN(m, 5), Array("�E�E�E"), tailN(m, 5))
        z = iota(2, m(UBound(m)) ^ 2)
            m = filterR(z, mapF(p_isPrime_(, m), z))
            printM catVs(headN(m, 5), Array("�E�E�E"), tailN(m, 5))
        z = iota(2, m(UBound(m)) ^ 2)
            m = filterR(z, mapF(p_isPrime_(, m), z))
            printM catVs(headN(m, 5), Array("�E�E�E"), tailN(m, 5))
    End Sub
    
    Private Sub newton_test()
        Dim m As Variant, z As Variant
        Debug.Print "------- �P����Newton�@�ɂ�鑽�����̍��i2�ʂ�j ------------"
        Debug.Print " f(x) = 2x^3 + x^2 - 5x + 0.5 �̗�_"
        z = Array(2, 1, -5, 0.5)
        m = VBA.Array(p_poly(, z), p_poly(, poly_deriv(z)))
        Debug.Print "x0 = 0  -> x = " & foldl_Funs(0, repeat(p_Newton_Raphson(, m), 7))
        Debug.Print "x0 = 1  -> x = " & repeat_while(1, p_less(0.0001, p_abs(m(0), 0)), p_Newton_Raphson(, m))
        Debug.Print "x0 = -7 -> x = " & repeat_while(-7, p_less(0.0001, p_abs(m(0), 0)), p_Newton_Raphson(, m))
    End Sub

    Private Sub find_test()
        Dim points As Variant, m As Variant, z As Variant
        Debug.Print "------- �����ɂ��Find ------------"
        Debug.Print "������ ( [0.0�`100.0] * 10000�� ) ���� 29.9�� 29.99�����̂��̂�T��"
        points = uniform_real_dist(10000, 0#, 100#)
        z = p_mult(p_greater(, 29.9), p_less(, 29.99))
        m = find_pred(z, points)
        If UBound(points) < m Then Debug.Print "�Ȃ�" Else Debug.Print points(m) & " (index=" & m & ")"
    End Sub

    ' ========== vbaUnit2 �̃e�X�g=================
    Private Sub flatten_test()
        Dim m As Variant, z As Variant
        Debug.Print "------- �z��̃t���b�g�� ------------"
        Debug.Print "m = makeM(3, 3, iota(1, 9))"
        m = makeM(3, 3, iota(1, 9))
        Debug.Print "m(1, 1) = m"
        m(1, 1) = m
        Debug.Print "m(1, 1)(1, 1) = m"
        m(1, 1)(1, 1) = m
        Debug.Print "--- m ---"
        printM m
        Debug.Print "--- m(1, 1) ---"
        printM m(1, 1)
        Debug.Print "--- m(1, 1)(1, 1) ---"
        printM m(1, 1)(1, 1)
        Debug.Print "--- m(1, 1)(1, 1)(1, 1) ---"
        printM m(1, 1)(1, 1)(1, 1)
        Debug.Print "--- flatten(m) ---"
        printM flatten(m)
    End Sub

    Private Sub sort_test2()
        Dim m As Variant:   ReDim m(0 To 9)
        Debug.Print "------�^���o���o���Ŕz����܂ޔz��̃\�[�g-----"
        m(0) = 75676786
        m(1) = "abc"
        m(2) = iota(1, 8)
        m(3) = "������"
        m(4) = 6
        m(5) = makeM(2, 2, iota(1, 4))
        m(6) = 300
        m(7) = iota(1, 15)
        m(8) = "������"
        m(9) = makeM(2, 3, iota(1, 6))
        Debug.Print "===�\�[�g�O==="
        Dim z As Variant
        For Each z In m
            Debug.Print "-+-+-+-+-+-"
            printM z
        Next z
        Debug.Print "===�\�[�g��==="
        m = subM(m, sortIndex_pred(m, p_comp000))
        For Each z In m
            Debug.Print "-+-+-+-+-+-"
            printM z
        Next z
    End Sub

    Private Sub sort_test3()
        Dim comp4 As Variant, arr As Variant, tmp As Variant, result As Variant
        Debug.Print "==== ������ёւ��ĉ\�ȍő吔��Ԃ��e�X�g ===="
        Debug.Print "1�`99 �̐���������20���"
        arr = uniform_int_dist(20, 0, 99)
        printM arr
        '--------------------
        comp4 = p_less(p_CLng(p_str_cat(ph_1, ph_2)), p_CLng(p_str_cat(ph_2, ph_1)))
        Debug.Print "��r�֐��͂��ꁫ  f(x, y) = CLng(x & y) < CLng(y & x)  �ɑ�������"
        Debug.Print "  p_less(p_CLng(p_str_cat(ph_1, ph_2)), p_CLng(p_str_cat(ph_2, ph_1)))"
        '--------------------
        tmp = mapF(p_CStr, arr)    ' ������
        Debug.Print "��r�֐��Ń\�[�g�i�t���j"
        result = subM(tmp, reverse(sortIndex_pred(tmp, comp4)))
        printM result
        Debug.Print "��"
        Debug.Print "  ";: Debug.Print foldl1(p_plus, result)     ' ��������������ĕ\��
    End Sub

    Private Sub equal_range_test()
        Debug.Print "==== equal_range�̃e�X�g ===="
        Dim m As Variant, erange As Variant
        m = uniform_int_dist(20, 1, 7)
        permutate m, sortIndex(m)
        printM catR(a_rows(m), m)
        printM "  ------------------------"
        erange = unzip(mapF_swap(p_equal_range, m, iota(1, 7)))
        Dim n As Long:  n = rowSize(erange(0))
        Dim tmp As Variant
        tmp = move_many(iota(1, n), repeat(":", n), repeat("[", n), erange(0), repeat("�`", n), erange(1), repeat(")", n))
        printM foldl1(p_catC, tmp)
    End Sub

    '�����̕���������̃��X�g�����o���u�֐��v���O���~���O���H����v�̖��
    Private Sub segmentsTest1()
        Debug.Print "==== �����̕���������̃��X�g�����o��Haskell�֐�(1) ===="
        Debug.Print "= segmentsTest1�i�w���p�֐��g�p�o�[�W�����j ="
        Debug.Print "    segments :: [a] -> [[a]]"
        Debug.Print "    segments = foldr (++) [] . scanr (\a b -> [a] : map (a:) b) []"
        Dim a As Variant, f As Variant, m As Variant
        a = Array("A", "B", "C", "D", "E")
        Debug.Print "�����foldr���遫"
        Debug.Print "scanr(p_cons(p_makeSole, p_consMap(ph_1, ph_2))"
        f = p_cons(p_makeSole, p_consMap(ph_1, ph_2))
        m = foldr(p_catV, Array(), scanr(f, Array(), a))
        Dim mm As Variant
        mm = scanr(f, Array(), a)
        mm = foldr(p_catV, Array(), mm)
        Debug.Print "Array(""A"", ""B"", ""C"", ""D"", ""E"") ��W�J����"
        printM mapF(p_join(, ""), mm)
    End Sub
    ' \a b -> [a] : map (a:) b �̂����Amap (a:) b �̕���
    Private Function consMap(ByRef a As Variant, ByRef v As Variant) As Variant
        consMap = mapF(p_cons(a), v)
    End Function
    Private Function p_consMap(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_consMap = make_funPointer(AddressOf consMap, firstParam, secondParam)
    End Function
    
    '�����̕���������̃��X�g�����o���u�֐��v���O���~���O���H����v�̖��
    '�A�h�z�b�N�Ȋ֐� consMap ���g�p���Ȃ��čςނ悤�ɂ���
    Private Sub segmentsTest2()
        Debug.Print "==== �����̕���������̃��X�g�����o��Haskell�֐�(2) ===="
        Debug.Print "    segments :: [a] -> [[a]]"
        Debug.Print "    segments = foldr (++) [] . scanr (\a b -> [a] : map (a:) b) []"
        Dim a As Variant, f As Variant, m As Variant
        a = Array("A", "B", "C", "D", "E")
        Debug.Print "�����foldr���遫"
        Debug.Print "scanr(p_cons(p_makeSole, p_mapF(p_cons(ph_1, yield_1), ph_2))"
        f = p_cons(p_makeSole, p_mapF(p_cons(ph_1, yield_1), ph_2))
        f = p_cons(p_makeSole, p_mapF(p_cons(ph_1, yield_1), ph_2))
        m = foldr(p_catV, Array(), scanr(f, Array(), a))
        Debug.Print "Array(""A"", ""B"", ""C"", ""D"", ""E"") ��W�J����"
        printM mapF(p_join(, ""), m)
    End Sub

    ' ����������ȍċA
    Private Sub curiouslyRecursiveTest()
        Dim arr As Variant
        Debug.Print "==== ����������ȍċA ===="
        Debug.Print " Array(1, Array(2, Array(3, Array(4, Array(5), 6))), 7) ��W�J"
        arr = Array(1, Array(2, Array(3, Array(4, Array(5), 6))), 7)
        Dim vec As Variant
        Set vec = curiouslyRecursive(New vh_stdvec, arr)
        vec.printS
        vec.printM
    End Sub
    Private Function curiouslyRecursive(ByRef vec As Variant, ByRef x As Variant) As Variant
        If IsArray(x) Then
            Set curiouslyRecursive = foldl(p_curiouslyRecursive, vec, x)
        Else
            Set curiouslyRecursive = vec.push_back(x)
        End If
    End Function
    Private Function p_curiouslyRecursive(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_curiouslyRecursive = make_funPointer(AddressOf curiouslyRecursive, firstParam, secondParam)
    End Function

' ==============================================================================
' ==============================================================================


    Private Function distance2(ByRef x As Variant, ByRef y As Variant) As Variant
        distance2 = ((x(0) - y(0)) ^ 2 + (x(1) - y(1)) ^ 2) ^ 0.5
    End Function
    Private Function p_distance2(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_distance2 = make_funPointer(AddressOf distance2, firstParam, secondParam)
    End Function

    '����
    Private Function u_real_rand(ByRef from_ As Variant, ByRef to_ As Variant) As Variant
        u_real_rand = uniform_real_dist(0, from_, to_)
    End Function
    Private Function p_rnd(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_rnd = make_funPointer(AddressOf u_real_rand, firstParam, secondParam)
    End Function

    '�t�B�{�i�b�`�֐�   (0,1)->(1,1)->(1,2)->(2,3)->(3,5)->(5,8)-> ...
    Private Function fibonacci(ByRef a As Variant, Optional ByRef dummy As Variant) As Variant
        fibonacci = Array(a(LBound(a) + 1), a(LBound(a)) + a(LBound(a) + 1))
    End Function
    Private Function p_fibonacci(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_fibonacci = make_funPointer(AddressOf fibonacci, firstParam, secondParam)
    End Function

    '���W�X�e�B�b�N�ʑ�   (Xn, r)->Xn+1
    Private Function Logistic(ByRef x As Variant, ByRef r As Variant) As Variant
        Logistic = r * x * (1 - x)
    End Function
    Private Function p_Logistic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_Logistic = make_funPointer(AddressOf Logistic, firstParam, secondParam)
    End Function

    '�~�^��
    Private Function circle_(ByRef point As Variant, ByRef r As Variant) As Variant
        circle_ = VBA.Array(point, r)
    End Function
    Private Function p_circle(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_circle = make_funPointer(AddressOf circle_, firstParam, secondParam)
    End Function

    '�~�̖ʐρ^���̑̐�
    Private Function circleArea(ByRef circle__ As Variant, ByRef dummy As Variant) As Variant
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
    Private Function p_circleArea(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_circleArea = make_funPointer(AddressOf circleArea, firstParam, secondParam)
    End Function

    '����inner product
    Private Function innerProduct_(ByRef a As Variant, ByRef b As Variant) As Variant
        innerProduct_ = foldl1(p_plus, zipWith(p_mult, a, b))
    End Function
    Private Function p_innerProduct_(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_innerProduct_ = make_funPointer(AddressOf innerProduct_, firstParam, secondParam)
    End Function

    '�s���
    Private Function matrixMult_(ByRef a As Variant, ByRef b As Variant) As Variant
        matrixMult_ = product_set(p_innerProduct_, zipC(a), zipR(b))
    End Function
    Function p_matrixMult_(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_matrixMult_ = make_funPointer(AddressOf matrixMult_, firstParam, secondParam)
    End Function

    '�f������(val�F����Ώۂ̎��R���Apm : �����̑f����, val��pm�̍ő吔��2��𒴂��Ȃ�����)
    Private Function isPrime_(ByRef val As Variant, ByRef pm As Variant) As Variant
        Dim z As Variant
        For Each z In pm
            If val < z * z Then Exit For
            If val Mod z = 0 Then
                isPrime_ = 0
                Exit Function
            End If
        Next z
        isPrime_ = 1
    End Function
    Private Function p_isPrime_(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_isPrime_ = make_funPointer(AddressOf isPrime_, firstParam, secondParam)
    End Function

    '�j���[�g���@�ɂ�鋁���̂P�X�e�b�v�@�F�@x1 ���� x2 ���o�͂���
    '��1���� �F�@x ,  ��2���� (f, df/dx)
    Private Function Newton_Raphson(ByRef x As Variant, ByRef fdf As Variant) As Variant
        Newton_Raphson = x - applyFun(x, fdf(0)) / applyFun(x, fdf(1))
    End Function
    Function p_Newton_Raphson(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_Newton_Raphson = make_funPointer(AddressOf Newton_Raphson, firstParam, secondParam)
    End Function

    '�������̔����i�W���̂݁j
    Private Function poly_deriv(ByRef coef As Variant) As Variant
        poly_deriv = headN(zipWith(p_mult, coef, iota(sizeof(coef) - 1, 0)), sizeof(coef) - 1)
    End Function

    '�������e���̐����̍��v�Ŕ�r����
    Private Function compareSS(ByRef a As Variant, ByRef b As Variant) As Variant
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
    Private Function p_compareSS(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_compareSS = make_funPointer(AddressOf compareSS, firstParam, secondParam)
    End Function

    '�^���o���o���Ŕz����܂ޔ�r�֐�
    Private Function comp000(ByRef a As Variant, ByRef b As Variant) As Variant
        ' �^���Ⴄ�ꍇ
        If VarType(a) <> VarType(b) Then
            comp000 = IIf(VarType(a) < VarType(b), 1, 0)
        Else
            ' �z��̏ꍇ
            If IsArray(a) Then
                ' �������قȂ�ꍇ
                If Dimension(a) <> Dimension(b) Then
                    comp000 = IIf(Dimension(a) < Dimension(b), 1, 0)
                Else
                    comp000 = IIf(sizeof(a) < sizeof(b), 1, 0)
                End If
            ElseIf IsNull(a) Then
                comp000 = 0
            Else    ' ���l�܂��͕�����
                comp000 = IIf(a < b, 1, 0)
            End If
        End If
    End Function
    Private Function p_comp000(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_comp000 = make_funPointer(AddressOf comp000, firstParam, secondParam)
    End Function


'========�؍\���̃e�X�g======================
'�^���o���o���Ŕz����܂ޖ؍\���̃e�X�g (���x�I�Ɏ��p���͖���)
Sub treeTest()
    Dim nodes As Variant, tree As Variant, n As Long, t As Long
    Dim DIC As Variant, i As Long
    Debug.Print "==== �^���o���o���Ŕz����܂ރL�[�ɂ��؍\���̃e�X�g ===="
    '===============�m�[�h�̏W��===============
    nodes = Array(makeNode(75676786, "A", p_comp000) _
                , makeNode("abc", "B", p_comp000) _
                , makeNode(iota(1, 8), "C", p_comp000) _
                , makeNode("������", "D", p_comp000) _
                , makeNode(6, "E", p_comp000) _
                , makeNode(makeM(2, 2, iota(1, 4)), "F", p_comp000) _
                , makeNode(300, "G", p_comp000) _
                , makeNode(iota(1, 15), "H", p_comp000) _
                , makeNode("������", "I", p_comp000) _
                , makeNode(makeM(2, 3, iota(1, 6)), "J", p_comp000) _
              )
    '===============��ݍ��݂ɂ��؂̍\�z===============
    tree = foldr(p_insertNode, Empty, nodes)
    '===============�L�[�̑I��===============
    Debug.Print """abc"" => ";
    printM getNode("abc", tree)
    Debug.Print """������"" => ";
    printM getNode("������", tree)
    Debug.Print "iota(1, 8) => ";
    printM getNode(iota(1, 8), tree)
    '==========================================================
    n = 10000
    Debug.Print "==== 0�`" & n & " �����_�������L�[ ===="
    nodes = zipWith(p_makeNode0, uniform_int_dist(n, 0, n), iota(1, n))
    t = GetTickCount
    tree = foldr(p_insertNode, Empty, nodes)
    Debug.Print GetTickCount - t & "ms"
    printM mapF(p_getNode(, tree), iota(0, 10))
    
    Set DIC = CreateObject("Scripting.Dictionary")
    t = GetTickCount
    For i = UBound(nodes) To LBound(nodes) Step -1
        DIC.Item(nodes(i)(0)) = nodes(i)(1)
    Next i
    Debug.Print GetTickCount - t & "ms  " & DIC.count & "_Items"
    For i = 0 To 10 Step 1
        Debug.Print DIC.Item(i);
    Next i
    Debug.Print ""
    Set DIC = Nothing
End Sub
    
    Private Function makeNode(ByRef key As Variant, ByRef val As Variant, Optional ByRef comp As Variant) As Variant
        Dim ret As Variant:     ReDim ret(0 To 4)
        ret(0) = key
        ret(1) = val
        ret(2) = Empty  'Left
        ret(3) = Empty  'Right
        ret(4) = IIf(IsMissing(comp), Empty, comp)
        makeNode = moveVariant(ret)
    End Function

    Private Function makeNode0(ByRef key As Variant, ByRef val As Variant) As Variant
        makeNode0 = makeNode(key, val)
    End Function
    Private Function p_makeNode0(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_makeNode0 = make_funPointer(AddressOf makeNode0, firstParam, secondParam)
    End Function

    Private Function less_with(ByRef a As Variant, ByRef b As Variant, ByRef comp As Variant) As Boolean
        If IsEmpty(comp) Then
            less_with = (a < b)
        Else
            less_with = 0 <> unbind_invoke(comp, a, b)
        End If
    End Function

    Private Function insertNode(ByRef node As Variant, ByRef tree As Variant) As Variant
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
    Private Function p_insertNode(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_insertNode = make_funPointer(AddressOf insertNode, firstParam, secondParam)
    End Function

    Private Function getNode(ByRef key As Variant, ByRef tree As Variant) As Variant
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
    Private Function p_getNode(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getNode = make_funPointer(AddressOf getNode, firstParam, secondParam)
    End Function

'C - �F�B�̗F�B
'http://abc016.contest.atcoder.jp/tasks/abc016_3
'�o�͗�
Sub test_friendsFriend()
    Dim inArr As Variant
    inArr = VBA.Array(VBA.Array(3, 2), _
                      VBA.Array(1, 2), _
                      VBA.Array(2, 3))
    printM friendsFriend(inArr)   '  1  0  1

    inArr = VBA.Array(VBA.Array(3, 3), _
                      VBA.Array(1, 2), _
                      VBA.Array(1, 3), _
                      VBA.Array(2, 3))
    printM friendsFriend(inArr)   '  0  0  0

    inArr = VBA.Array(VBA.Array(8, 12), _
                      VBA.Array(1, 6), _
                      VBA.Array(1, 7), _
                      VBA.Array(1, 8), _
                      VBA.Array(2, 5), _
                      VBA.Array(2, 6), _
                      VBA.Array(3, 5), _
                      VBA.Array(3, 6), _
                      VBA.Array(4, 5), _
                      VBA.Array(4, 8), _
                      VBA.Array(5, 6), _
                      VBA.Array(5, 7), _
                      VBA.Array(7, 8))
    printM friendsFriend(inArr)   '  4  4  4  5  2  3  4  2
End Sub

    '�F�B�̗F�B�֐�
    Private Function friendsFriend(ByRef inArr As Variant)
        Dim fMatrix As Variant
        '�F�B�}�g���N�X(ID��1�n�܂肾���z��C���f�b�N�X�Ƃ���0�n�܂�ɕύX)
        fMatrix = makeM(inArr(0)(0), inArr(0)(0), 0)
        Dim i As Long
        For i = LBound(inArr) + 1 To UBound(inArr) Step 1
            fMatrix(inArr(i)(0) - 1, inArr(i)(1) - 1) = 1
            fMatrix(inArr(i)(1) - 1, inArr(i)(0) - 1) = 1
        Next i
        '�e���[�U�̗F�B�W��
        Dim myFriends As Variant
        myFriends = mapF(p_rowMax, mapF(p_filterR(fMatrix), zipR(fMatrix)))
        '���ڂ̗F�B�����O
        myFriends = zipWith(p_sumOfLess, zipR(fMatrix), myFriends)
        '�������g�����O
        friendsFriend = mapF(p_minus(, 1), myFriends)
    End Function

    '2�����s��̍s�����̍ő�l
    Private Function rowMax(ByRef matrix As Variant, ByRef dummy As Variant) As Variant
        rowMax = foldl1(p_max, matrix, 1)
    End Function
        Function p_rowMax(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
            p_rowMax = make_funPointer(AddressOf rowMax, firstParam, secondParam)
        End Function
    
    '�ӂ���1�����s��a,b�̊e�v�f�ɂ���(a�̒l < b�̒l)�̌�
    Private Function sumOfLess(ByRef a As Variant, ByRef b As Variant) As Variant
        sumOfLess = foldl1(p_plus, zipWith(p_less, a, b))
    End Function
        Function p_sumOfLess(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
            p_sumOfLess = make_funPointer(AddressOf sumOfLess, firstParam, secondParam)
        End Function
