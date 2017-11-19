Attribute VB_Name = "misc_ratio"
'misc_ratio
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'**************************************************************
'�@���{������ bigInt
'--------------------------------------------------------------
'   Function long2bigInt            Long����bigInt�𐶐�
'   Function bigInt2double          bigInt����Double�𐶐��i�I�[�o�[�t���[�΍�Ȃ��j
'   Function double2bigInt          Double����bigInt�𐶐��i�s���m�j
'   Function bigInt_log             bigInt�̑ΐ�
'   Function bigInt_sgn             bigInt�̕���
'   Function bigInt_base            bigInt�̊
'   Function bigInt_abs             bigInt�̐�Βl
'   Function bigInt_plus            bigInt�̉��Z
'   Function bigInt_minus           bigInt�̌��Z
'   Function bigInt_mult            bigInt�̏�Z
'   Function bigInt_divide_mod      bigInt�̏��Z (����Mod)
'   Function bigInt_divide          bigInt�̏��Z
'   Function bigInt_mod             bigInt��Mod
'   Function bigInt_pow             bigInt�̃x�L��
'   Function bigInt2str             bigInt����String�ւ̕ϊ��i10�i�\���j
'   Function str2bigInt             String����bigInt�ւ̕ϊ�
'   Function bigInt_equal           bigInt�̔�r (a = b)
'   Function bigInt_not_equal       bigInt�̔�r (a <> b)
'   Function bigInt_less            bigInt�̔�r (a < b)
'   Function bigInt_less_equal      bigInt�̔�r (a <= b)
'   Function bigInt_greater         bigInt�̔�r (a > b)
'   Function bigInt_greater_equal   bigInt�̔�r (a >= b)
'   Function bigInt_min             bigInt��min
'   Function bigInt_max             bigInt��max
'   Function bigInt_gcd             �ő����
'**************************************************************
Private Const int_15 As Long = 2 ^ 15

    ' bigInt�̐��K��
    Private Function bigInt_normalize(ByRef bigInt As Variant, _
                            ByVal shorten As Boolean, _
                        ByVal baseN As Long) As Variant
        Dim ret As Variant
        Dim UB As Long: UB = UBound(bigInt)
        Dim flg As Boolean:     flg = False
        If shorten Then
            Do While 0 < UB And 0 = bigInt(UB): UB = UB - 1:    Loop
        Else
            If 0 = bigInt(UB) Then UB = UB - 1
        End If
        ReDim ret(0 To UB + 1)
        Dim i As Long
        For i = 1 To UB Step 1
            If 0 < bigInt(i) Then flg = True
            ret(i) = ret(i) + bigInt(i)
            ret(i + 1) = ret(i) \ baseN
            ret(i) = ret(i) Mod baseN
        Next i
        ret(0) = IIf(flg, bigInt(0), 0)
        swapVariant ret, bigInt_normalize
    End Function

' Long ���� bigInt�𐶐�
    Private Function long2bigInt_imple(ByVal num As Long, ByVal baseN As Long) As Variant
        Dim valN As Long:    valN = Abs(num)
        long2bigInt_imple = bigInt_normalize( _
                    VBA.Array(Sgn(num) * baseN, valN Mod baseN, valN \ baseN, 0), _
                    True, baseN)
    End Function
Function long2bigInt(ByRef num As Variant, Optional ByRef dummy As Variant) As Variant
    long2bigInt = long2bigInt_imple(num, int_15)
End Function
    Function p_long2bigInt(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_long2bigInt = make_funPointer(AddressOf long2bigInt, firstParam, secondParam)
    End Function

' bigInt ���� Double�𐶐��i�I�[�o�[�t���[�΍�Ȃ��j
Function bigInt2double(ByRef bigInt As Variant, Optional ByRef dummy As Variant) As Variant
    Dim ret As Double, i As Long
    Dim baseN As Long:  baseN = bigInt_base(bigInt)
    ret = 0#
    If 0 <> bigInt_sgn(bigInt) Then
        For i = UBound(bigInt) To 1 Step -1
            ret = ret * baseN + bigInt(i)
        Next i
    End If
    bigInt2double = bigInt_sgn(bigInt) * ret
End Function
    Function p_bigInt2double(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt2double = make_funPointer(AddressOf bigInt2double, firstParam, secondParam)
    End Function

' Double ���� bigInt�𐶐��i�s���m�j
Function double2bigInt(ByRef dval As Variant, Optional ByRef dummy As Variant) As Variant
    If Abs(dval) < 1 Then
        double2bigInt = long2bigInt(0)
    Else
        Dim dval2 As Double:    dval2 = Abs(dval)
        Dim N As Long:          N = Fix(Log(dval2) / Log(int_15))
        Dim ret As Variant, i As Long
        ReDim ret(0 To N + 1)
        For i = N + 1 To 1 Step -1
            ret(i) = Fix(dval2 * int_15 ^ (-i + 1))
            dval2 = dval2 - ret(i) * int_15 ^ (i - 1)
        Next i
        ret(0) = Sgn(dval) * int_15
        swapVariant ret, double2bigInt
    End If
End Function
    Function p_double2bigInt(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_double2bigInt = make_funPointer(AddressOf double2bigInt, firstParam, secondParam)
    End Function

' bigInt�̑ΐ�
Function bigInt_log(ByRef bigInt As Variant, Optional ByRef dummy As Variant) As Variant
    Dim baseN As Long:  baseN = bigInt_base(bigInt)
    Dim ret As Double
    ret = -1.01E-14
    If 0 <> bigInt_sgn(bigInt) Then
        Dim i As Long
        For i = 1 To UBound(bigInt) Step 1
            ret = ret / baseN + bigInt(i)
        Next i
    End If
    bigInt_log = (UBound(bigInt) - 1) * Log(baseN) + Log(ret)
End Function
    Function p_bigInt_log(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_log = make_funPointer(AddressOf bigInt_log, firstParam, secondParam)
    End Function

' bigInt�̕���
Function bigInt_sgn(ByRef bigInt As Variant, Optional ByRef dummy As Variant) As Variant
    bigInt_sgn = Sgn(bigInt(0))
End Function
    Function p_bigInt_sgn(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_sgn = make_funPointer(AddressOf bigInt_sgn, firstParam, secondParam)
    End Function

' bigInt�̊
Function bigInt_base(ByRef bigInt As Variant, Optional ByRef dummy As Variant) As Variant
    bigInt_base = Abs(bigInt(0))
End Function
    Function p_bigInt_base(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_base = make_funPointer(AddressOf bigInt_base, firstParam, secondParam)
    End Function

' bigInt�̐�Βl
Function bigInt_abs(ByRef bigInt As Variant, Optional ByRef dummy As Variant) As Variant
    Dim ret As Variant
    ret = bigInt
    ret(0) = bigInt_base(ret)
    swapVariant bigInt_abs, ret
End Function
    Function p_bigInt_abs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_abs = make_funPointer(AddressOf bigInt_abs, firstParam, secondParam)
    End Function

    Private Function bigInt_end_pos(ByRef bigInt As Variant) As Long
        For bigInt_end_pos = UBound(bigInt) To 0 Step -1
            If 0 < bigInt(bigInt_end_pos) Then Exit For
        Next bigInt_end_pos
    End Function
    '
    Private Function bigInt_abs_less(ByRef leftV As Variant, ByRef rightV As Variant) As Variant
        Dim dif As Long:    dif = bigInt_end_pos(rightV) - bigInt_end_pos(leftV)
        If 0 < dif Then
            bigInt_abs_less = 1
        ElseIf dif < 0 Then
            bigInt_abs_less = 0
        Else
            Dim i As Long
            bigInt_abs_less = 0
            For i = bigInt_end_pos(leftV) To 1 Step -1
                If leftV(i) < rightV(i) Then
                    bigInt_abs_less = 1:    Exit For
                ElseIf rightV(i) < leftV(i) Then
                    bigInt_abs_less = 0:    Exit For
                End If
            Next i
        End If
    End Function
        
    ' bigInt�̃x�[�X�ϊ�
    Private Function bigInt_base_change(ByRef bigInt As Variant, ByVal baseN As Long) As Variant
        Dim baseO As Long:  baseO = bigInt_base(bigInt)
        If baseO = baseN Then
            bigInt_base_change = bigInt
        Else
            Dim pos1 As Variant, posN As Variant, ret As Variant, tmp As Variant
            Dim i As Long, N As Long
            N = bigInt_end_pos(bigInt)
            pos1 = long2bigInt_imple(baseO, baseN)
            posN = pos1
            ret = long2bigInt_imple(bigInt(1), baseN)
            For i = 2 To N Step 1
                tmp = long2bigInt_imple(bigInt(i), baseN)
                ret = bigInt_plus(ret, bigInt_mult(tmp, posN))
                posN = bigInt_mult(pos1, posN)
            Next i
            ret(0) = ret(0) * bigInt_sgn(bigInt)
            swapVariant bigInt_base_change, ret
        End If
    End Function
    
Function bigInt_plus(ByRef bigInt1 As Variant, ByRef bigInt2 As Variant) As Variant
    If IsNumeric(bigInt1) Then
        If VarType(bigInt1) = vbDouble Then
            bigInt_plus = bigInt_plus_inple1(double2bigInt(bigInt1), bigInt2)
        Else
            bigInt_plus = bigInt_plus_inple1(long2bigInt(bigInt1), bigInt2)
        End If
    Else
        bigInt_plus = bigInt_plus_inple1(bigInt1, bigInt2)
    End If
End Function
    Function p_bigInt_plus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_plus = make_funPointer(AddressOf bigInt_plus, firstParam, secondParam)
    End Function
    
    Private Function bigInt_plus_inple1(ByRef bigInt1 As Variant, ByRef bigInt2 As Variant) As Variant
        If IsNumeric(bigInt2) Then
            If VarType(bigInt2) = vbDouble Then
                bigInt_plus_inple1 = bigInt_plus_inple2(bigInt1, double2bigInt(bigInt2))
            Else
                bigInt_plus_inple1 = bigInt_plus_inple2(bigInt1, long2bigInt(bigInt2))
            End If
        Else
            bigInt_plus_inple1 = bigInt_plus_inple2(bigInt1, bigInt2)
        End If
    End Function
    
    Private Function bigInt_plus_inple2(ByRef bigInt1 As Variant, ByRef bigInt2 As Variant) As Variant
        Dim ret As Variant, i As Long
        Dim baseN As Long
        baseN = bigInt_base(bigInt1):   If baseN = 0 Then baseN = bigInt_base(bigInt2)
        If bigInt_sgn(bigInt1) = 0 Then
            bigInt_plus_inple2 = bigInt2
        ElseIf bigInt_sgn(bigInt2) = 0 Then
            bigInt_plus_inple2 = bigInt1
        ElseIf bigInt_sgn(bigInt1) = bigInt_sgn(bigInt2) Then
            ReDim ret(0 To 1 + max_fun(bigInt_end_pos(bigInt1), bigInt_end_pos(bigInt2)))
            For i = 1 To bigInt_end_pos(bigInt1) Step 1
                ret(i) = bigInt1(i)
            Next i
            For i = 1 To bigInt_end_pos(bigInt2) Step 1
                ret(i) = ret(i) + bigInt2(i)
            Next i
            ret(0) = bigInt1(0)
            bigInt_plus_inple2 = bigInt_normalize(ret, True, baseN)
        ElseIf bigInt_abs_less(bigInt1, bigInt2) = 1 Then
            ReDim ret(0 To 1 + max_fun(bigInt_end_pos(bigInt1), bigInt_end_pos(bigInt2)))
            For i = 1 To bigInt_end_pos(bigInt2) Step 1
                ret(i) = bigInt2(i)
            Next i
            For i = 1 To bigInt_end_pos(bigInt1) Step 1
                ret(i) = ret(i) - bigInt1(i)
                If ret(i) < 0 Then
                    ret(i) = ret(i) + baseN
                    ret(i + 1) = ret(i + 1) - 1
                End If
            Next i
            ret(0) = bigInt2(0)
            bigInt_plus_inple2 = bigInt_normalize(ret, True, baseN)
        Else
            ReDim ret(0 To 1 + max_fun(bigInt_end_pos(bigInt1), bigInt_end_pos(bigInt2)))
            For i = 1 To bigInt_end_pos(bigInt1) Step 1
                ret(i) = bigInt1(i)
            Next i
            For i = 1 To bigInt_end_pos(bigInt2) Step 1
                ret(i) = ret(i) - bigInt2(i)
                If ret(i) < 0 Then
                    ret(i) = ret(i) + baseN
                    ret(i + 1) = ret(i + 1) - 1
                End If
            Next i
            For i = i To UBound(ret) - 1 Step 1
                If ret(i) < 0 Then
                    ret(i) = ret(i) + baseN
                    ret(i + 1) = ret(i + 1) - 1
                End If
            Next i
            ret(0) = bigInt1(0)
            bigInt_plus_inple2 = bigInt_normalize(ret, True, baseN)
        End If
    End Function

' bigInt�̌��Z
Function bigInt_minus(ByRef bigInt1 As Variant, ByRef bigInt2 As Variant) As Variant
    Dim tmp2 As Variant
    If IsNumeric(bigInt2) Then
        If VarType(bigInt2) = vbDouble Then
            tmp2 = double2bigInt(bigInt2)
        Else
            tmp2 = long2bigInt(bigInt2)
        End If
    Else
        tmp2 = bigInt2
    End If
    tmp2(0) = -tmp2(0)
    bigInt_minus = bigInt_plus(bigInt1, tmp2)
End Function
    Function p_bigInt_minus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_minus = make_funPointer(AddressOf bigInt_minus, firstParam, secondParam)
    End Function

' bigInt�̏�Z
Function bigInt_mult(ByRef bigInt1 As Variant, ByRef bigInt2 As Variant) As Variant
    If IsNumeric(bigInt1) Then
        If VarType(bigInt1) = vbDouble Then
            bigInt_mult = bigInt_mult_imple1(double2bigInt(bigInt1), bigInt2)
        Else
            bigInt_mult = bigInt_mult_imple1(long2bigInt(bigInt1), bigInt2)
        End If
    Else
        bigInt_mult = bigInt_mult_imple1(bigInt1, bigInt2)
    End If
End Function
    Function p_bigInt_mult(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_mult = make_funPointer(AddressOf bigInt_mult, firstParam, secondParam)
    End Function

    Private Function bigInt_mult_imple1(ByRef bigInt1 As Variant, ByRef bigInt2 As Variant) As Variant
    If IsNumeric(bigInt2) Then
        If VarType(bigInt2) = vbDouble Then
            bigInt_mult_imple1 = bigInt_mult_imple2(bigInt1, double2bigInt(bigInt2))
        Else
            bigInt_mult_imple1 = bigInt_mult_imple2(bigInt1, long2bigInt(bigInt2))
        End If
    Else
        bigInt_mult_imple1 = bigInt_mult_imple2(bigInt1, bigInt2)
    End If
    End Function

    Private Function bigInt_mult_imple2(ByRef bigInt1 As Variant, ByRef bigInt2 As Variant) As Variant
        If UBound(bigInt2) < UBound(bigInt1) Then
            bigInt_mult_imple2 = bigInt_mult_imple2(bigInt2, bigInt1)
        Else
            Dim ret As Variant, i As Long, j As Long
            Dim baseN As Long:  baseN = bigInt_base(bigInt1)
            Dim sb As Long: sb = bigInt_sgn(bigInt1) * bigInt_sgn(bigInt2) * bigInt_base(bigInt1)
            ReDim ret(0 To UBound(bigInt1) + UBound(bigInt2) - 1)
            For i = 1 To UBound(bigInt1) Step 1
                ret(0) = sb
                If sb = 0 Then Exit For
                For j = 1 To UBound(bigInt2) Step 1
                    ret(i + j - 1) = ret(i + j - 1) + bigInt1(i) * bigInt2(j)
                Next j
                ret = bigInt_normalize(ret, False, baseN)
            Next i
            bigInt_mult_imple2 = bigInt_normalize(ret, True, baseN)
        End If
    End Function


    ' Log�̒l ���� bigInt�̍ő區�𐶐�
    Private Function log2bigInt(ByVal dval As Double, ByVal baseN As Long) As Variant
        Dim N As Long
        If dval < 0 Then
            log2bigInt = long2bigInt(0)
        Else
            Dim logintN As Double: logintN = Log(baseN)
            N = Fix(dval / logintN)
            Dim ret As Variant
            ret = repeat(0, N + 2)
            ret(0) = baseN
            ret(N + 1) = Fix(Exp(dval - N * logintN))
            swapVariant log2bigInt, ret
        End If
    End Function
    
' bigInt�̏��Z�i����Mod�j
Function bigInt_divide_mod(ByRef bigIntT As Variant, ByRef bigIntB As Variant) As Variant
    If IsNumeric(bigIntT) Then
        If VarType(bigIntT) = vbDouble Then
            bigInt_divide_mod = bigInt_divide_mod_imple1(double2bigInt(bigIntT), bigIntB)
        Else
            bigInt_divide_mod = bigInt_divide_mod_imple1(long2bigInt(bigIntT), bigIntB)
        End If
    Else
        bigInt_divide_mod = bigInt_divide_mod_imple1(bigIntT, bigIntB)
    End If
End Function
    Function p_bigInt_divide_mod(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_divide_mod = make_funPointer(AddressOf bigInt_divide_mod, firstParam, secondParam)
    End Function

    Private Function bigInt_divide_mod_imple1(ByRef bigIntT As Variant, ByRef bigIntB As Variant) As Variant
    If IsNumeric(bigIntB) Then
        If VarType(bigIntB) = vbDouble Then
            bigInt_divide_mod_imple1 = bigInt_divide_mod_imple2(bigIntT, double2bigInt(bigIntB))
        Else
            bigInt_divide_mod_imple1 = bigInt_divide_mod_imple2(bigIntT, long2bigInt(bigIntB))
        End If
    Else
        bigInt_divide_mod_imple1 = bigInt_divide_mod_imple2(bigIntT, bigIntB)
    End If
    End Function

    Private Function bigInt_divide_mod_imple2(ByRef bigIntT As Variant, ByRef bigIntB As Variant) As Variant
        Dim logR As Double
        Dim copyT As Variant:   copyT = bigInt_abs(bigIntT)
        Dim copyB As Variant:   copyB = bigInt_abs(bigIntB)
        Dim baseN As Long:  baseN = bigInt_base(bigIntB)
        Dim logB As Double: logB = bigInt_log(copyB, baseN)
        logR = bigInt_log(copyT) - logB
        Dim div As Variant
        div = log2bigInt(logR, baseN)
        Dim modd As Variant: modd = copyT
        Do While Not bigInt_abs_less(modd, copyB)
            modd = bigInt_minus(copyT, bigInt_mult(copyB, div))
            If bigInt_sgn(modd) = 0 Then Exit Do
            If bigInt_sgn(modd) < 0 Then
                div = bigInt_minus(div, 1)
                modd = bigInt_minus(copyT, bigInt_mult(copyB, div))
                If bigInt_sgn(modd) = 0 Then Exit Do
            End If
            logR = bigInt_log(modd) - logB
            If logR < 0 Then Exit Do
            div = bigInt_plus(div, log2bigInt(logR, baseN))
        Loop
        Dim ret As Variant: ReDim ret(0 To 1)
        swapVariant ret(0), div
        swapVariant ret(1), modd
        swapVariant bigInt_divide_mod_imple2, ret
    End Function

' bigInt�̏��Z
Function bigInt_divide(ByRef bigIntT As Variant, ByRef bigIntB As Variant) As Variant
    bigInt_divide = bigInt_divide_mod(bigIntT, bigIntB)(0)
End Function
    Function p_bigInt_divide(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_divide = make_funPointer(AddressOf bigInt_divide, firstParam, secondParam)
    End Function

' bigInt��Mod
Function bigInt_mod(ByRef bigIntT As Variant, ByRef bigIntB As Variant) As Variant
    bigInt_mod = bigInt_divide_mod(bigIntT, bigIntB)(1)
End Function
    Function p_bigInt_mod(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_mod = make_funPointer(AddressOf bigInt_mod, firstParam, secondParam)
    End Function

' bigInt�̃x�L��
Function bigInt_pow(ByRef bigInt As Variant, ByRef nv As Variant) As Variant
    Dim xx As Variant
    If IsNumeric(bigInt) Then
        If VarType(bigInt) = vbDouble Then
            xx = double2bigInt(bigInt)
        Else
            xx = long2bigInt(bigInt)
        End If
    Else
        xx = bigInt
    End If
    Dim ret As Variant:     ret = long2bigInt_imple(1, bigInt_base(xx))
    Dim N As Long:          N = nv
    Do While 0 < N
        If (1 = N Mod 2) Then ret = bigInt_mult(ret, xx)
        xx = bigInt_mult(xx, xx)
        N = N \ 2
    Loop
    swapVariant bigInt_pow, ret
End Function
    Function p_bigInt_pow(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_pow = make_funPointer(AddressOf bigInt_pow, firstParam, secondParam)
    End Function

' bigInt����String�ւ̕ϊ��i10�i�\���j
Function bigInt2str(ByRef bigInt As Variant, Optional ByRef dummy As Variant) As Variant
    If IsNumeric(bigInt) Then
        bigInt2str = str(bigInt)
    ElseIf bigInt_base(bigInt) = 10000 Then
        Dim N As Long:  N = bigInt_end_pos(bigInt)
        Dim i As Long
        Dim ret As String
        ret = IIf(0 <= bigInt_sgn(bigInt), "", "-")
        ret = ret & bigInt(N)
        For i = N - 1 To 1 Step -1
            ret = ret & right("0000" & bigInt(i), 4)
        Next i
        bigInt2str = ret
    ElseIf bigInt_sgn(bigInt) = 0 Then
        bigInt2str = "0"
    Else
        Dim tmp As Variant
        tmp = bigInt2str(bigInt_base_change(bigInt, 10000))
        swapVariant bigInt2str, tmp
    End If
End Function
    Function p_bigInt2str(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt2str = make_funPointer(AddressOf bigInt2str, firstParam, secondParam)
    End Function

' String����bigInt�ւ̕ϊ�
Function str2bigInt(ByRef bstr As Variant, Optional ByRef dummy As Variant) As Variant
    Dim sign As Long:   sign = 1
    Dim baseN As Long:  baseN = 10
    Dim spos As Long:   spos = 1
    Dim pref As String: pref = ""
    If left(bstr, 1) = "-" Then
        sign = -1
        spos = 2
    End If
    If Mid(bstr, spos, 1) = 0 Then
        baseN = 8
        spos = spos + 1
        pref = "&O"
        If StrConv(Mid(bstr, spos, 1), vbNarrow + vbLowerCase) = "x" Then
            baseN = 16
            spos = spos + 1
            pref = "&H"
        End If
    End If
    Dim i As Long, tmp As Long
    Dim ret As Variant
    Dim lpos As Long:   lpos = 1
    ReDim ret(0 To 1 + Fix((Len(bstr) - spos + 1) / 2))
    For i = Len(bstr) - 1 To spos Step -2
        ret(lpos) = CLng(pref & Mid(bstr, i, 2))
        lpos = lpos + 1
    Next i
    If spos - 2 < i Then
        ret(lpos) = CLng(pref & Mid(bstr, spos, i + 2 - spos))
    End If
    ret(0) = sign * baseN * baseN
    str2bigInt = bigInt_base_change(bigInt_normalize(ret, True, baseN * baseN), int_15)
End Function
    Function p_str2bigInt(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_str2bigInt = make_funPointer(AddressOf str2bigInt, firstParam, secondParam)
    End Function

' bigInt�̔�r  (a = b)
Function bigInt_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    bigInt_equal = IIf(bigInt_less(a, b) Or bigInt_less(b, a), 0, 1)
End Function
    Function p_bigInt_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_equal = make_funPointer(AddressOf bigInt_equal, firstParam, secondParam)
    End Function

' bigInt�̔�r  (a <> b)
Function bigInt_not_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    bigInt_not_equal = IIf(bigInt_equal(a, b), 0, 1)
End Function
    Function p_bigInt_not_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_not_equal = make_funPointer(AddressOf bigInt_not_equal, firstParam, secondParam)
    End Function

' bigInt�̔�r  (a < b)
Function bigInt_less(ByRef a As Variant, ByRef b As Variant) As Variant
    If IsNumeric(a) Then
        If IsNumeric(b) Then
            bigInt_less = (a < b)
        ElseIf VarType(a) = vbDouble Then
            bigInt_less = bigInt_less_imple1(double2bigInt(a), b)
        Else
            bigInt_less = bigInt_less_imple1(long2bigInt(a), b)
        End If
    Else
        bigInt_less = bigInt_less_imple1(a, b)
    End If
End Function
    Function p_bigInt_less(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_less = make_funPointer(AddressOf bigInt_less, firstParam, secondParam)
    End Function

    Private Function bigInt_less_imple1(ByRef a As Variant, ByRef b As Variant) As Variant
    If IsNumeric(b) Then
        If VarType(b) = vbDouble Then
            bigInt_less_imple1 = bigInt_less_imple2(a, double2bigInt(b))
        Else
            bigInt_less_imple1 = bigInt_less_imple2(a, long2bigInt(b))
        End If
    Else
        bigInt_less_imple1 = bigInt_less_imple2(a, b)
    End If
    End Function

    Private Function bigInt_less_imple2(ByRef a As Variant, ByRef b As Variant) As Variant
        If bigInt_sgn(b) < bigInt_sgn(a) Then
            bigInt_less_imple2 = 0
        ElseIf bigInt_sgn(a) < bigInt_sgn(b) Then
            bigInt_less_imple2 = 1
        ElseIf bigInt_sgn(a) = 0 Then
            bigInt_less_imple2 = 0
        ElseIf 0 < bigInt_sgn(a) Then
            bigInt_less_imple2 = bigInt_abs_less(a, b)
        Else
            bigInt_less_imple2 = bigInt_abs_less(b, a)
        End If
    End Function

' bigInt�̔�r  (a <= b)
Function bigInt_less_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    bigInt_less_equal = IIf(bigInt_less(b, a), 0, 1)
End Function
    Function p_bigInt_less_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_less_equal = make_funPointer(AddressOf bigInt_less_equal, firstParam, secondParam)
    End Function

' bigInt�̔�r  (a > b)
Function bigInt_greater(ByRef a As Variant, ByRef b As Variant) As Variant
    bigInt_greater = bigInt_less(b, a)
End Function
    Function p_bigInt_greater(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_greater = make_funPointer(AddressOf bigInt_greater, firstParam, secondParam)
    End Function

' bigInt�̔�r  (a >= b)
Function bigInt_greater_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    bigInt_greater_equal = IIf(bigInt_less(a, b), 0, 1)
End Function
    Function p_bigInt_greater_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_greater_equal = make_funPointer(AddressOf bigInt_greater_equal, firstParam, secondParam)
    End Function

' bigInt��min
Function bigInt_min(ByRef a As Variant, ByRef b As Variant) As Variant
    bigInt_min = IIf(bigInt_less(a, b), a, b)
End Function
    Function p_bigInt_min(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_min = make_funPointer(AddressOf bigInt_min, firstParam, secondParam)
    End Function

' bigInt��max
Function bigInt_max(ByRef a As Variant, ByRef b As Variant) As Variant
    bigInt_max = IIf(bigInt_less(a, b), b, a)
End Function
    Function p_bigInt_max(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_max = make_funPointer(AddressOf bigInt_max, firstParam, secondParam)
    End Function

'�ő����
Function bigInt_gcd(ByRef a As Variant, ByRef b As Variant) As Variant
    If IsNumeric(a) Then
        If VarType(a) = vbDouble Then
            bigInt_gcd = bigInt_gcd_imple1(double2bigInt(a), b)
        Else
            bigInt_gcd = bigInt_gcd_imple1(long2bigInt(a), b)
        End If
    Else
        bigInt_gcd = bigInt_gcd_imple1(a, b)
    End If
End Function
    Function p_bigInt_gcd(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_bigInt_gcd = make_funPointer(AddressOf bigInt_gcd, firstParam, secondParam)
    End Function

    Private Function bigInt_gcd_imple1(ByRef a As Variant, ByRef b As Variant) As Variant
    If IsNumeric(b) Then
        If VarType(b) = vbDouble Then
            bigInt_gcd_imple1 = bigInt_gcd_imple2(a, double2bigInt(b))
        Else
            bigInt_gcd_imple1 = bigInt_gcd_imple2(a, long2bigInt(b))
        End If
    Else
        bigInt_gcd_imple1 = bigInt_gcd_imple2(a, b)
    End If
    End Function
    
    Private Function bigInt_gcd_imple2(ByRef a As Variant, ByRef b As Variant) As Variant
        If bigInt_sgn(a) = 0 Then
            bigInt_gcd_imple2 = long2bigInt(1)
        ElseIf bigInt_sgn(b) = 0 Then
            bigInt_gcd_imple2 = bigInt_abs(a)
        Else
            bigInt_gcd_imple2 = bigInt_gcd(b, bigInt_mod(bigInt_abs(a), bigInt_abs(b)))
        End If
    End Function


'**************************************************************
'�@�L�����̉��Z
'--------------------------------------------------------------
'�@�L�����̃f�[�^�\���̓N���X���������̒P���Ȕz��Ƃ���
'  Array(���q, ����) �F ���q As Long, ���� As Long�i�񕉁j
'  �I�[�o�[�t���[�ɂ��đ΍�͂��Ă��邪�K�[�h�Ȃ�
'  �[�����Z�ɂ��Ă̓K�[�h�Ȃ�
'--------------------------------------------------------------
'   Function make_ratio         :   �L�����̐���
'   Function ratio2double       :   Double�ɕϊ�
'   Function ratio2str          :   String�ɕϊ�
'   Function ratio_plus         :   �L�����̉��Z
'   Function ratio_negate       :   �L�����̕����ύX
'   Function ratio_minus        :   �L�����̌��Z
'   Function ratio_mult         :   �L�����̏�Z
'   Function ratio_pow          :   �L�����̃x�L��
'   Function ratio_divide       :   �L�����̏��Z
'   Function ratio_sgn          :   �L�����̕���
'   Function ratio_abs          :   �L�����̐�Βl
'   Function ratio_equal        :   �L�����̔�r  (a = b)
'   Function ratio_not_equal    :   �L�����̔�r  (a <> b)
'   Function ratio_less         :   �L�����̔�r  (a < b)
'   Function ratio_less_equal   :   �L�����̔�r  (a <= b)
'   Function ratio_greater      :   �L�����̔�r  (a > b)
'   Function ratio_greater_equal:   �L�����̔�r  (a >= b)
'**************************************************************

    '�ő����
    Public Function getGcd(ByVal a As Long, ByVal b As Long) As Long
        If a = 0 Then
            getGcd = 1
        ElseIf b = 0 Then
            getGcd = Abs(a)
        Else
            getGcd = getGcd(b, Abs(a) Mod Abs(b))
        End If
    End Function

'�L�����̐���
Function make_ratio(ByRef num As Variant, ByRef den As Variant) As Variant
    Dim gcd As Long:    gcd = getGcd(num, den)
    make_ratio = VBA.Array(Sgn(num * den) * (Abs(num) \ gcd), Abs(den) \ gcd)
End Function
    Function p_make_ratio(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_make_ratio = make_funPointer(AddressOf make_ratio, firstParam, secondParam)
    End Function

'Double�ɕϊ�
Function ratio2double(ByRef ratio As Variant, Optional ByRef secondParam As Variant) As Variant
    ratio2double = ratio(0) / ratio(1)
End Function
    Function p_ratio2double(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio2double = make_funPointer(AddressOf ratio2double, firstParam, secondParam)
    End Function

'String�ɕϊ�
Function ratio2str(ByRef ratio As Variant, Optional ByRef secondParam As Variant) As Variant
    ratio2str = CStr(ratio(0)) & "/" & CStr(ratio(1))
End Function
    Function p_ratio2str(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio2str = make_funPointer(AddressOf ratio2str, firstParam, secondParam)
    End Function

'�L�����̉��Z
Function ratio_plus(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
        Dim n1 As Long: n1 = ratio_1(0)
        Dim d1 As Long: d1 = ratio_1(1)
        Dim n2 As Long: n2 = ratio_2(0)
        Dim d2 As Long: d2 = ratio_2(1)
        Dim gcd As Long:    gcd = getGcd(d1, d2)
    ratio_plus = make_ratio(n1 * (d2 \ gcd) + n2 * (d1 \ gcd), d1 * (d2 \ gcd))
End Function
    Function p_ratio_plus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_plus = make_funPointer(AddressOf ratio_plus, firstParam, secondParam)
    End Function

'�L�����̕����ύX
Function ratio_negate(ByRef ratio As Variant, Optional ByRef dummy As Variant) As Variant
    ratio_negate = make_ratio(-ratio(0), ratio(1))
End Function
    Function p_ratio_negate(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_negate = make_funPointer(AddressOf ratio_negate, firstParam, secondParam)
    End Function

'�L�����̌��Z
Function ratio_minus(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
    ratio_minus = ratio_plus(ratio_1, ratio_negate(ratio_2))
End Function
    Function p_ratio_minus(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_minus = make_funPointer(AddressOf ratio_minus, firstParam, secondParam)
    End Function

'�L�����̏�Z
Function ratio_mult(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
        Dim n1 As Long: n1 = ratio_1(0)
        Dim d1 As Long: d1 = ratio_1(1)
        Dim n2 As Long: n2 = ratio_2(0)
        Dim d2 As Long: d2 = ratio_2(1)
        Dim gx As Long:    gx = getGcd(n1, d2)
        Dim gy As Long:    gy = getGcd(n2, d1)
    ratio_mult = make_ratio((n1 \ gx) * (n2 \ gy), (d2 \ gx) * (d1 \ gy))
End Function
    Function p_ratio_mult(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_mult = make_funPointer(AddressOf ratio_mult, firstParam, secondParam)
    End Function

'�L�����̃x�L��
Function ratio_pow(ByRef ratio As Variant, ByRef N As Variant) As Variant
        Dim n1 As Long: n1 = ratio(0)
        Dim d1 As Long: d1 = ratio(1)
    If N = 0 Then
        ratio_pow = make_ratio(1, 1)
    ElseIf 0 < N Then
        ratio_pow = make_ratio(n1 ^ N, d1 ^ N)
    Else
        ratio_pow = make_ratio(d1 ^ -N, n1 ^ -N)
    End If
End Function
    Function p_ratio_pow(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_pow = make_funPointer(AddressOf ratio_pow, firstParam, secondParam)
    End Function

'�L�����̏��Z
Function ratio_divide(ByRef ratio_1 As Variant, ByRef ratio_2 As Variant) As Variant
        Dim n2 As Long: n2 = ratio_2(0)
        Dim d2 As Long: d2 = ratio_2(1)
    ratio_divide = ratio_mult(ratio_1, make_ratio(d2, n2))
End Function
    Function p_ratio_divide(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_divide = make_funPointer(AddressOf ratio_divide, firstParam, secondParam)
    End Function

'�L�����̕���
Function ratio_sgn(ByRef ratio As Variant, Optional ByRef dummy As Variant) As Variant
    ratio_sgn = Sgn(ratio(0))
End Function
    Function p_ratio_sgn(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_sgn = make_funPointer(AddressOf ratio_sgn, firstParam, secondParam)
    End Function

'�L�����̐�Βl
Function ratio_abs(ByRef ratio As Variant, Optional ByRef dummy As Variant) As Variant
    ratio_abs = make_ratio(Abs(ratio(0)), ratio(1))
End Function
    Function p_ratio_abs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_abs = make_funPointer(AddressOf ratio_abs, firstParam, secondParam)
    End Function

'�L�����̔�r  (a = b)
Function ratio_equal(ByRef a As Variant, ByRef b As Variant) As Variant
        Dim n1 As Long: n1 = a(0)
        Dim d1 As Long: d1 = a(1)
        Dim n2 As Long: n2 = b(0)
        Dim d2 As Long: d2 = b(1)
    ratio_equal = IIf(n1 = n2 And d1 = d2, 1, 0)
End Function
    Function p_ratio_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_equal = make_funPointer(AddressOf ratio_equal, firstParam, secondParam)
    End Function

'�L�����̔�r  (a <> b)
Function ratio_not_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    ratio_not_equal = IIf(ratio_equal(a, b) = 1, 0, 1)
End Function
    Function p_ratio_not_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_not_equal = make_funPointer(AddressOf ratio_not_equal, firstParam, secondParam)
    End Function

'�L�����̔�r  (a < b)
Function ratio_less(ByRef a As Variant, ByRef b As Variant) As Variant
    ratio_less = IIf(ratio2double(a) < ratio2double(b), 1, 0)
End Function
    Function p_ratio_less(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_less = make_funPointer(AddressOf ratio_less, firstParam, secondParam)
    End Function

'�L�����̔�r  (a <= b)
Function ratio_less_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    ratio_less_equal = IIf(ratio_less(b, a), 0, 1)
End Function
    Function p_ratio_less_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_less_equal = make_funPointer(AddressOf ratio_less_equal, firstParam, secondParam)
    End Function

'�L�����̔�r  (a > b)
Function ratio_greater(ByRef a As Variant, ByRef b As Variant) As Variant
    ratio_greater = ratio_less(b, a)
End Function
    Function p_ratio_greater(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_greater = make_funPointer(AddressOf ratio_greater, firstParam, secondParam)
    End Function

'�L�����̔�r  (a >= b)
Function ratio_greater_equal(ByRef a As Variant, ByRef b As Variant) As Variant
    ratio_greater_equal = IIf(ratio_less(a, b), 0, 1)
End Function
    Function p_ratio_greater_equal(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_ratio_greater_equal = make_funPointer(AddressOf ratio_greater_equal, firstParam, secondParam)
    End Function


