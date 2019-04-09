Attribute VB_Name = "Haskell_1_Core"
'Haskell_1_Core
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'***********************************************************************************
'   �֐��^�v���O����
' API��Callback�Ƃ��ēn����֐��̃V�O�l�`����
' Function fun(ByRef x As Variant, ByRef y As Variant) As Variant
' ��������
' Function fun(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
'===================================================================================
'   Function moveVariant        source��VARIANT�ϐ���target��VARIANT��move����
'   Function ph_0               �v���[�X�z���_
'   Function ph_1               �v���[�X�z���_
'   Function ph_2               �v���[�X�z���_
'   Function yield_0            �֐��]������ ph_0 �𐶐�����
'   Function yield_1            �֐��]������ ph_1 �𐶐�����
'   Function yield_2            �֐��]������ ph_2 �𐶐�����
'   Function make_funPointer    ���[�U�֐���bind�t�@���N�^������i�֐��̕����K�p�j
'   Function make_funPointer_with_2nd_Default  2�Ԗڂ̈����Ƀf�t�H���g�l��ݒ肷��ꍇ
'   Function is_bindFun         bind���ꂽ�֐��ł��邱�Ƃ̔���
'   Function bind1st            1�Ԗڂ̈������đ�������
'   Function bind2nd            2�Ԗڂ̈������đ�������
'   Sub      swap1st            ��1������swap�i�傫�ȕϐ��̏ꍇ�Ɏg�p�j
'   Sub      swap2nd            ��2������swap�i�傫�ȕϐ��̏ꍇ�Ɏg�p�j
' * Function mapF               �z��̊e�v�f�Ɋ֐���K�p����
'   Function mapF_swap          mapF(fun(a), m) �܂��� mapF(fun(, a), m)�̍\����
'   Function applyFun           �֐��K�p�֐�
'   Function setParam           �֐��Ɉ�������
'   Function foldl_Funs         �֐������ifoldl�j
'   Function scanl_Funs         �֐������iscanl�j
'   Function foldr_Funs         �֐������ifoldr�j
'   Function scanr_Funs         �֐������iscanr�j
'   Function applyFun2by2       ((x, y), (f1, f2, ...)) -> Array(f1(x, y), f2(x, y), ...)
'   Function setParam2by2       ((f1, f2, ...), (x, y)) -> Array(f1(x, y), f2(x, y), ...)
'   Function count_if           �z��̊e�v�f�ŏq��ɂ��]�����ʂ��[���łȂ����̂̐�
'   Function find_pred          1�����z�񂩂�����ɍ��v������̂�����
'   Function find_best_pred     1�����z�񂩂�����ɍŗǍ��v������̂�����
'   Function repeat_while       �q��ɂ����������������ԌJ��Ԃ��֐��K�p
'   Function repeat_while_not   �q��ɂ���������������Ȃ��ԌJ��Ԃ��֐��K�p
'   Function generate_while     �q��ɂ����������������ԌJ��Ԃ��֐��K�p�̗����𐶐�
'   Function generate_while_not �q��ɂ���������������Ȃ��ԌJ��Ԃ��֐��K�p�̗����𐶐�
'   Function p_foldl            1�����z������foldl
'   Function p_foldr            1�����z������foldr
'   Function p_foldl1           1�����z���foldl1
'   Function p_foldr1           1�����z���foldr1
'   Function p_scanl            1�����z������scanl
'   Function p_scanr            1�����z������scanr
'   Function p_scanl1           1�����z���scanl1
'   Function p_scanr1           1�����z���scanr1
'   Function foldl_zipWith      zipWith��foldl����
'   Function foldl1_zipWith     zipWith��foldl1����
'   Function foldr_zipWith      zipWith��foldr����
'   Function foldr1_zipWith     zipWith��foldr1����
'   Function scanl_zipWith      zipWith��scanl����
'   Function scanr_zipWith      zipWith��scanr����
'   Function scanl1_zipWith     zipWith��scanl1����
'   Function scanr1_zipWith     zipWith��scanr1����
'***********************************************************************************

'source��VARIANT�ϐ���target��VARIANT��move����
Function moveVariant(ByRef source As Variant, Optional ByRef dummy As Variant) As Variant
    swapVariant moveVariant, source
End Function
    Function p_moveVariant(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_moveVariant = make_funPointer(AddressOf moveVariant, firstParam, secondParam)
    End Function
'***********************************************************************************

'�v���[�X�z���_�i�u���ꂽ�ʒu�ɂ���đ�1�����������͑�2�������󂯎��j
Function ph_0() As Variant
    ph_0 = placeholder(0)
End Function

'�v���[�X�z���_�i��1�������󂯎��j
Function ph_1() As Variant
    ph_1 = placeholder(1)
End Function

'�v���[�X�z���_�i��2�������󂯎��j
Function ph_2() As Variant
    ph_2 = placeholder(2)
End Function

'�֐��]������ ph_0 �𐶐�����
Function yield_0() As Variant
    yield_0 = placeholder(800)
End Function

'�֐��]������ ph_1 �𐶐�����
Function yield_1() As Variant
    yield_1 = placeholder(801)
End Function

'�֐��]������ ph_2 �𐶐�����
Function yield_2() As Variant
    yield_2 = placeholder(802)
End Function

    ' Array() �� IsMissing = True �ɂȂ邱�Ƃ�WorkAround
    Function Is_Missing_(Optional ByRef x As Variant) As Boolean
        Is_Missing_ = IIf(IsMissing(x) And Not IsArray(x), True, False)
    End Function

'���[�U�֐���bind�t�@���N�^������i�֐��̕����K�p�j
'make_funPointer(func)                              �����̑����Ȃ�
'make_funPointer(func, firstParam)                  1�Ԗڂ̈����𑩔�
'make_funPointer(func, , secondParam)               2�Ԗڂ̈����𑩔�
'make_funPointer(func, firstParam, secondParam)     �����̈����𑩔��i�x���]���j
'functionParamPoint = 1 : firstParam���֐�
'functionParamPoint = 2 : secondParam���֐�
'functionParamPoint = 0 : firstParam�AsecondParam���l�i�f�t�H���g�j
Function make_funPointer(ByVal func As LongPtr, _
                         ByRef firstParam As Variant, _
                         ByRef secondParam As Variant, _
                         Optional ByVal functionParamPoint As Long = 0) As Variant
    make_funPointer = VBA.Array(func, _
                    IIf(Is_Missing_(firstParam), yield_0, firstParam), _
                    IIf(Is_Missing_(secondParam), yield_0, secondParam), _
                    placeholder(functionParamPoint) _
                   )
End Function

'���[�U�֐���bind�t�@���N�^������i2�Ԗڂ̈����Ƀf�t�H���g�l��ݒ肷��ꍇ�j
Function make_funPointer_with_2nd_Default(ByVal func As LongPtr, _
                         ByRef firstParam As Variant, _
                         ByRef secondParam As Variant, _
                         Optional ByVal functionParamPoint As Long = 0) As Variant
    make_funPointer_with_2nd_Default = VBA.Array(func, _
                                 IIf(Is_Missing_(firstParam), yield_0, firstParam), _
                                 secondParam, _
                                 placeholder(functionParamPoint) _
                                )
End Function

'bind���ꂽ�֐��ł��邱�Ƃ̔���
Function is_bindFun(ByRef val As Variant) As Boolean
    is_bindFun = False
    If Dimension(val) = 1 And sizeof(val) = 4 Then is_bindFun = is_placeholder(val(3))
End Function

'�������đ�������
    Private Function bind_imple(ByRef func As Variant, _
                                ByRef param As Variant, _
                                ByVal pA As Long, _
                                ByVal pB As Long, _
                                ByVal pC As Long, _
                                ByVal pD As Long) As Variant
        If is_bindFun(func) Then
            bind_imple = VBA.Array(func(0), _
                                   bind_imple(func(1), param, pA, pB, pC, pD), _
                                   bind_imple(func(2), param, pA, pB, pC, pD), _
                                   func(3))
        ElseIf is_placeholder(func) Then
            If func = placeholder(pA) Or func = placeholder(pB) Or func = placeholder(pC) Or func = placeholder(pD) Then
                bind_imple = param
            Else
                bind_imple = func
            End If
        Else
            bind_imple = func
        End If
    End Function

Function bind1st(ByRef func As Variant, ByRef firstParam As Variant, _
                    Optional ByVal ph_only As Boolean = False) As Variant
    If ph_only Then
        bind1st = VBA.Array(func(0), _
                            bind_imple(func(1), firstParam, 0, 1, 0, 1), _
                            bind_imple(func(2), firstParam, 1, 1, 1, 1), _
                            func(3))
    Else
        bind1st = VBA.Array(func(0), _
                            bind_imple(func(1), firstParam, 0, 1, 800, 801), _
                            bind_imple(func(2), firstParam, 1, 1, 801, 801), _
                            func(3))
    End If
End Function

Function bind2nd(ByRef func As Variant, ByRef secondParam As Variant, _
                    Optional ByVal ph_only As Boolean = False) As Variant
    If ph_only Then
        bind2nd = VBA.Array(func(0), _
                            bind_imple(func(1), secondParam, 2, 2, 2, 2), _
                            bind_imple(func(2), secondParam, 0, 0, 2, 2), _
                            func(3))
    Else
        bind2nd = VBA.Array(func(0), _
                            bind_imple(func(1), secondParam, 2, 2, 802, 802), _
                            bind_imple(func(2), secondParam, 0, 2, 800, 802), _
                            func(3))
    End If
End Function

'��1������swap�i�傫�ȕϐ��̏ꍇ�Ɏg�p�j
Sub swap1st(ByRef func As Variant, ByRef firstParam As Variant)
    If is_bindFun(func) Then swapVariant func(1), firstParam
End Sub

'��2������swap�i�傫�ȕϐ��̏ꍇ�Ɏg�p�j
Sub swap2nd(ByRef func As Variant, ByRef secondParam As Variant)
    If is_bindFun(func) Then swapVariant func(2), secondParam
End Sub

'*************************************************************************
' �z��̊e�v�f�Ɋ֐���K�p����
Function mapF(ByRef func As Variant, ByRef matrix As Variant) As Variant
    mapF = mapF_imple(func, matrix)
End Function
    Function p_mapF(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_mapF = make_funPointer(AddressOf mapF, firstParam, secondParam, 1)
    End Function

' mapF(fun(a), m) �܂��� mapF(fun(, a), m)�̍\����
' ������a��move����A������߂����B�傫�Ȕz��̃R�s�[����������B
' �p�^�[��1  mapF_swap(fun, a, m)      -> mapF(fun(a), m)
' �p�^�[��2  mapF_swap(fun, , b, m)    -> mapF(fun(, b), m)
' �p�^�[��3  mapF_swap(fun, a, , m)    -> mapF(fun(a), m)    �i�p�^�[��1�Ɠ����j
' �p�^�[��4  mapF_swap(fun, a, b, m)   -> mapF(fun(a, b), m) �i�֎~�j
Function mapF_swap(ByRef fun As Variant, _
                        Optional ByRef x As Variant, _
                        Optional ByRef y As Variant, _
                        Optional ByRef z As Variant) As Variant
    If Is_Missing_(z) Then          ' �p�^�[��1
        swap1st fun, x
        mapF_swap = mapF(fun, y)
        swap1st fun, x
    Else
        If Is_Missing_(x) Then      ' �p�^�[��2
            swap2nd fun, y
            mapF_swap = mapF(fun, z)
            swap2nd fun, y
        ElseIf Is_Missing_(y) Then  ' �p�^�[��3
            swap1st fun, x
            mapF_swap = mapF(fun, z)
            swap1st fun, x
        Else                        ' �p�^�[��4
            '
        End If
    End If
End Function


'*************************************************************************
'�֐��K�p�֐�  1�����ɑ΂��Ċ֐���K�p����   �֐���Bind��
'1. applyFun(x     ,  Null          )     ->  x
'2. applyFun(x     ,  Empty         )     ->  x
'3. applyFun(x     , (f, a, b)      )     ->  f(a, b)
'4. applyFun(x     , (f, a) )             ->  f(a, x)
'5. applyFun(x     , (f, , b) )           ->  f(x, b)
Function applyFun(ByRef param As Variant, ByRef func As Variant) As Variant
    If IsNull(func) Or IsEmpty(func) Then
        applyFun = param
    Else
        applyFun = unbind_invoke(func, param, param)
    End If
End Function
    Function p_applyFun(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_applyFun = make_funPointer(AddressOf applyFun, firstParam, secondParam, 2)
    End Function

'�֐���1������������֐�
'1. setParam(f              , x     )  ->  f(x)
'2. setParam((f, a, placeholder), x )  ->  f(a, x)
'3. setParam((f, placeholder, b), x )  ->  f(x, b)
Function setParam(ByRef func As Variant, ByRef param As Variant) As Variant
    setParam = applyFun(param, func)
End Function
    Function p_setParam(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_setParam = make_funPointer(AddressOf setParam, firstParam, secondParam, 1)
    End Function

'�֐������ifoldl�j
Function foldl_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    foldl_Funs = foldl(p_applyFun, init, funcArray)
End Function
    Function p_foldl_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_foldl_Funs = make_funPointer(AddressOf foldl_Funs, firstParam, secondParam)
    End Function

'�֐������iscanl�j
Function scanl_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    scanl_Funs = scanl(p_applyFun, init, funcArray)
End Function
    Function p_scanl_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_scanl_Funs = make_funPointer(AddressOf scanl_Funs, firstParam, secondParam)
    End Function

'�֐������ifoldr�j
Function foldr_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    foldr_Funs = foldr(p_setParam, init, funcArray)
End Function
    Function p_foldr_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_foldr_Funs = make_funPointer(AddressOf foldr_Funs, firstParam, secondParam)
    End Function

'�֐������iscanr�j
Function scanr_Funs(ByRef init As Variant, ByRef funcArray As Variant) As Variant
    scanr_Funs = scanr(p_setParam, init, funcArray)
End Function
    Function p_scanr_Funs(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_scanr_Funs = make_funPointer(AddressOf scanr_Funs, firstParam, secondParam)
    End Function

'((x, y), f)  �ɑ΂���  f(x, y)     ��Ԃ�
'((x, y), (f1, f2, ...))  �ɑ΂���  Array(f1(x, y), f2(x, y), ...)     ��Ԃ�
Function applyFun2by2(ByRef params As Variant, ByRef funcs As Variant) As Variant
    Dim ret As Variant, z As Variant, k As Long: k = 0
    If is_bindFun(funcs) Then
         applyFun2by2 = unbind_invoke(funcs, params(LBound(params)), params(1 + LBound(params)))
    Else
        ReDim ret(0 To sizeof(funcs) - 1)
        For Each z In funcs
            ret(k) = unbind_invoke(z, params(LBound(params)), params(1 + LBound(params)))
            k = k + 1
        Next z
        swapVariant applyFun2by2, ret
    End If
End Function
    Function p_applyFun2by2(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_applyFun2by2 = make_funPointer(AddressOf applyFun2by2, firstParam, secondParam)
    End Function

'(f, (x, y))  �ɑ΂���  f(x, y)     ��Ԃ�
'((f1, f2, ...), (x, y))  �ɑ΂���  Array(f1(x, y), f2(x, y), ...)     ��Ԃ�
Function setParam2by2(ByRef funcs As Variant, ByRef params As Variant) As Variant
    setParam2by2 = applyFun2by2(params, funcs)
End Function
    Function p_setParam2by2(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_setParam2by2 = make_funPointer(AddressOf setParam2by2, firstParam, secondParam)
    End Function

' �z�� matrix �̊e�v�f�ŏq��ɂ��]�����ʂ��[���łȂ����̂̐�
Function count_if(ByRef pred As Variant, ByRef matrix As Variant) As Variant
    Dim z As Variant
    count_if = 0&
    For Each z In mapF(pred, matrix)
        If z <> 0 Then count_if = count_if + 1
    Next z
End Function
    Function p_count_if(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_count_if = make_funPointer(AddressOf count_if, firstParam, secondParam, 1)
    End Function

'1�����z�񂩂�����ɍ��v������̂�����(�ŏ��Ƀq�b�g�����C���f�b�N�X��Ԃ�)
'1�����z��ȊO�ł���ΕԂ�l��Empty�A���������ꍇ�� UBound + 1 ��Ԃ�
Function find_pred(ByRef pred As Variant, ByRef vec As Variant) As Variant
    If Dimension(vec) = 1 Then
        find_pred = find_imple(pred, vec, UBound(vec) + 1)
    End If
End Function
    Function p_find_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_find_pred = make_funPointer(AddressOf find_pred, firstParam, secondParam, 1)
    End Function

'1�����z�񂩂�����ɍŗǍ��v������̂�����(�ŏ��Ƀq�b�g�����C���f�b�N�X��Ԃ�)
'1�����z��ȊO�ł���ΕԂ�l��Empty
Function find_best_pred(ByRef pred As Variant, ByRef vec As Variant) As Variant
    If Dimension(vec) = 1 Then
        find_best_pred = find_best_imple(pred, vec, UBound(vec) + 1)
        If find_best_pred = UBound(vec) + 1 Then find_best_pred = Empty
    End If
End Function
    Function p_find_best_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_find_best_pred = make_funPointer(AddressOf find_best_pred, firstParam, secondParam, 1)
    End Function

' �q��ɂ����������������ԌJ��Ԃ��֐��K�p
Function repeat_while(ByRef val As Variant, _
                      ByRef pred As Variant, _
                      ByRef fun As Variant, _
                      Optional ByVal n As Long = -1) As Variant
    repeat_while = repeat_imple(val, pred, fun, n, 0, 0)
End Function

' �q��ɂ���������������Ȃ��ԌJ��Ԃ��֐��K�p
Function repeat_while_not(ByRef val As Variant, _
                          ByRef pred As Variant, _
                          ByRef fun As Variant, _
                          Optional ByVal n As Long = -1) As Variant
    repeat_while_not = repeat_imple(val, pred, fun, n, 0, 1)
End Function

' �q��ɂ����������������ԌJ��Ԃ��֐��K�p�̗����𐶐�
Function generate_while(ByVal val As Variant, _
                        ByRef pred As Variant, _
                        ByRef fun As Variant, _
                        Optional ByVal n As Long = -1) As Variant
    generate_while = repeat_imple(val, pred, fun, n, 1, 0)
End Function

' �q��ɂ���������������Ȃ��ԌJ��Ԃ��֐��K�p�̗����𐶐�
Function generate_while_not(ByVal val As Variant, _
                            ByRef pred As Variant, _
                            ByRef fun As Variant, _
                            Optional ByVal n As Long = -1) As Variant
    generate_while_not = repeat_imple(val, pred, fun, n, 1, 1)
End Function

' 1�����z������ foldl (p_foldl �̂�Public)
    Private Function foldl_v(ByRef fun_init As Variant, ByRef vec As Variant) As Variant
        Dim fun As Variant
        fun = bind2nd(bind1st(fun_init(0), vec, True), vec, True)
        foldl_v = foldl(fun, fun_init(1), vec)
    End Function
Public Function p_foldl(ByRef fun As Variant, ByRef init As Variant) As Variant
    p_foldl = VBA.Array(AddressOf foldl_v, _
                        VBA.Array(fun, init), _
                        yield_0, _
                        placeholder)
End Function

' 1�����z������ foldr (p_foldr �̂�Public)
    Private Function foldr_v(ByRef fun_init As Variant, ByRef vec As Variant) As Variant
        Dim fun As Variant
        fun = bind2nd(bind1st(fun_init(0), vec, True), vec, True)
        foldr_v = foldr(fun, fun_init(1), vec)
    End Function
Public Function p_foldr(ByRef fun As Variant, ByRef init As Variant) As Variant
    p_foldr = VBA.Array(AddressOf foldr_v, _
                        VBA.Array(fun, init), _
                        yield_0, _
                        placeholder)
End Function

' 1�����z������ foldl1 (p_foldl1 �̂�Public)
    Private Function foldl1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        foldl1_v = foldl1(fun, vec)
    End Function
Public Function p_foldl1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_foldl1 = make_funPointer(AddressOf foldl1_v, firstParam, secondParam, 1)
End Function

' 1�����z������ foldr1 (p_foldr1 �̂�Public)
    Private Function foldr1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        foldr1_v = foldr1(fun, vec)
    End Function
Public Function p_foldr1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_foldr1 = make_funPointer(AddressOf foldr1_v, firstParam, secondParam, 1)
End Function

' 1�����z������ scanl (p_scanl �̂�Public)
    Private Function scanl_v(ByRef fun_init As Variant, ByRef vec As Variant) As Variant
        Dim fun As Variant
        fun = bind2nd(bind1st(fun_init(0), vec, True), vec, True)
        scanl_v = scanl(fun, fun_init(1), vec)
    End Function
Public Function p_scanl(ByRef fun As Variant, ByRef init As Variant) As Variant
    p_scanl = VBA.Array(AddressOf scanl_v, _
                        VBA.Array(fun, init), _
                        yield_0, _
                        placeholder)
End Function

' 1�����z������ scanr (p_scanr �̂�Public)
    Private Function scanr_v(ByRef fun_init As Variant, ByRef vec As Variant) As Variant
        Dim fun As Variant
        fun = bind2nd(bind1st(fun_init(0), vec, True), vec, True)
        scanr_v = scanr(fun, fun_init(1), vec)
    End Function
Public Function p_scanr(ByRef fun As Variant, ByRef init As Variant) As Variant
    p_scanr = VBA.Array(AddressOf scanr_v, _
                        VBA.Array(fun, init), _
                        yield_0, _
                        placeholder)
End Function

' 1�����z������ scanl1 (p_scanl1 �̂�Public)
    Private Function scanl1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        scanl1_v = scanl1(fun, vec)
    End Function
Public Function p_scanl1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_scanl1 = make_funPointer(AddressOf scanl1_v, firstParam, secondParam, 1)
End Function

' 1�����z������ scanr1 (p_scanr1 �̂�Public)
    Private Function scanr1_v(ByRef fun As Variant, ByRef vec As Variant) As Variant
        scanr1_v = scanr1(fun, vec)
    End Function
Public Function p_scanr1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_scanr1 = make_funPointer(AddressOf scanr1_v, firstParam, secondParam, 1)
End Function

' zipWith��foldl����
Function foldl_zipWith(ByRef fun As Variant, ByRef init As Variant, ByRef vec As Variant) As Variant
    foldl_zipWith = init
    If 0 < sizeof(vec) Then
        Dim i As Long
        For i = LBound(vec) To UBound(vec) Step 1
            foldl_zipWith = zipWith(fun, foldl_zipWith, vec(i))
        Next i
    End If
End Function

' zipWith��foldr����
Function foldr_zipWith(ByRef fun As Variant, ByRef init As Variant, ByRef vec As Variant) As Variant
    foldr_zipWith = init
    If 0 < sizeof(vec) Then
        Dim i As Long
        For i = UBound(vec) To LBound(vec) Step -1
            foldr_zipWith = zipWith(fun, vec(i), foldr_zipWith)
        Next i
    End If
End Function

' zipWith��foldl1����
Function foldl1_zipWith(ByRef fun As Variant, ByRef vec As Variant) As Variant
    If 0 < sizeof(vec) Then
        foldl1_zipWith = vec(LBound(vec))
        Dim i As Long
        For i = LBound(vec) + 1 To UBound(vec) Step 1
            foldl1_zipWith = zipWith(fun, foldl1_zipWith, vec(i))
        Next i
    End If
End Function

' zipWith��foldr1����
Function foldr1_zipWith(ByRef fun As Variant, ByRef vec As Variant) As Variant
    If 0 < sizeof(vec) Then
        foldr1_zipWith = vec(UBound(vec))
        Dim i As Long
        For i = UBound(vec) - 1 To LBound(vec) Step -1
            foldr1_zipWith = zipWith(fun, vec(i), foldr1_zipWith)
        Next i
    End If
End Function

' zipWith��scanl����
Function scanl_zipWith(ByRef fun As Variant, ByRef init As Variant, ByRef vec As Variant) As Variant
    Dim ret As Variant: ret = makeM(1 + sizeof(vec))
    Dim i As Long, k As Long: k = 0
    ret(k) = init
    For i = LBound(vec) To UBound(vec) Step 1
        k = k + 1
        ret(k) = zipWith(fun, ret(k - 1), vec(i))
    Next i
    scanl_zipWith = moveVariant(ret)
End Function

' zipWith��scanr����
Function scanr_zipWith(ByRef fun As Variant, ByRef init As Variant, ByRef vec As Variant) As Variant
    Dim ret As Variant: ret = makeM(1 + sizeof(vec))
    Dim i As Long, k As Long: k = UBound(ret)
    ret(k) = init
    For i = UBound(vec) To LBound(vec) Step -1
        k = k - 1
        ret(k) = zipWith(fun, vec(i), ret(k + 1))
    Next i
    scanr_zipWith = moveVariant(ret)
End Function

' zipWith��scanl1����
Function scanl1_zipWith(ByRef fun As Variant, ByRef vec As Variant) As Variant
    Dim ret As Variant: ret = makeM(sizeof(vec))
    Dim i As Long, k As Long: k = 0
    ret(k) = vec(LBound(vec))
    For i = LBound(vec) + 1 To UBound(vec) Step 1
        k = k + 1
        ret(k) = zipWith(fun, ret(k - 1), vec(i))
    Next i
    scanl1_zipWith = moveVariant(ret)
End Function

' zipWith��scanr1����
Function scanr1_zipWith(ByRef fun As Variant, ByRef vec As Variant) As Variant
    Dim ret As Variant: ret = makeM(sizeof(vec))
    Dim i As Long, k As Long: k = UBound(ret)
    ret(k) = vec(UBound(vec))
    For i = UBound(vec) - 1 To LBound(vec) Step -1
        k = k - 1
        ret(k) = zipWith(fun, vec(i), ret(k + 1))
    Next i
    scanr1_zipWith = moveVariant(ret)
End Function
