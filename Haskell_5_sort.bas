Attribute VB_Name = "Haskell_5_sort"
'Haskell_5_sort
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'====================================================================================================
    ' Function  sortIndex           �����\�[�g��̃C���f�b�N�X�z��
    ' Function  sortIndex_pred      �C�ӂ̔�r�֐��ɂ��\�[�g��̃C���f�b�N�X�z��
    ' Sub       permutate           1�����z��̕��׊���
    ' Sub       permutate_back      permutate �ŕ��׊�����ꂽ1�����z������̏���ɖ߂�
    ' Function  lower_bound         �\�[�g�ςݔz�񂩂�̃L�[�̌����istd::lower_bound�Ɠ����j
    ' Function  lower_bound_pred    �\�[�g�ςݔz�񂩂�̃L�[�̌����istd::lower_bound�Ɠ����j
    ' Function  upper_bound         �\�[�g�ςݔz�񂩂�̃L�[�̌����istd::upper_bound�Ɠ����j
    ' Function  upper_bound_pred    �\�[�g�ςݔz�񂩂�̃L�[�̌����istd::upper_bound�Ɠ����j
    ' Function  equal_range         �\�[�g�ςݔz�񂩂�̃L�[�̌����istd::equal_range�Ɠ����j
    ' Function  equal_range_pred    �\�[�g�ςݔz�񂩂�̃L�[�̌����istd::equal_range�Ɠ����j
    ' Function  partition_points    �\�[�g�ςݔz�񂩂�����ɂ���ċ敪������Ă���ʒu�̈ꗗ�𓾂�
    ' Function  partition_points_pred
    ' Function  less_dic            �q�� ������less
    ' Function  less_equal_dic      �q�� ������less_equal
    ' Function  greater_dic         �q�� ������greater
    ' Function  greater_equal_dic   �q�� ������greater_equal
    ' Function  equal_dic           �q�� ������equal
    ' Function  notEqual_dic        �q�� ������notEqual
'====================================================================================================

'�����\�[�g��̃C���f�b�N�X�z��i�~���\�[�g�͂���reverse���Ƃ�j
'key_columns ��2�����z��̏ꍇ�̃L�[��w�� Array(0,2,4)
' �Ώ۔z������ۂɃ\�[�g����ꍇ�́Apermutate(�z��, sortIndex) �Ƃ��邩�A
' ��������subV(�z��, sortIndex) �����
Function sortIndex(ByRef matrix As Variant, Optional ByRef key_columns As Variant) As Variant
    Select Case Dimension(matrix)
    Case 1
        sortIndex = stdsort(matrix, 1, 0)
    Case 2
        Dim allkeyFlag As Boolean:  allkeyFlag = False
        If IsMissing(key_columns) Then
            key_columns = a_cols(matrix)
            allkeyFlag = True
        End If
        If sizeof(key_columns) = 1 Then
            sortIndex = stdsort(selectCol(matrix, key_columns(LBound(key_columns))), 1, 0)
        ElseIf allkeyFlag Then
            sortIndex = stdsort(zipC(matrix), 2, 0)
        Else
            sortIndex = stdsort(zipC(subM(matrix, , key_columns)), 2, 0)
        End If
    End Select
End Function
    Public Function p_sortIndex(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_sortIndex = make_funPointer_with_2nd_Default(AddressOf sortIndex, firstParam, secondParam)
    End Function

'�C�ӂ̔�r�֐� comp �ɂ��\�[�g��̃C���f�b�N�X�z��
Function sortIndex_pred(ByRef matrix As Variant, ByRef comp As Variant) As Variant
    Select Case Dimension(matrix)
    Case 1
        sortIndex_pred = stdsort(matrix, 0, comp)
    Case 2
        sortIndex_pred = stdsort(zipC(matrix), 0, comp)
    End Select
End Function
    Public Function p_sortIndex_pred(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_sortIndex_pred = make_funPointer(AddressOf sortIndex_pred, firstParam, secondParam, 2)
    End Function

' 1�����z�� vec �̕��׊���
' s_index �� sortIndex �֐��A�������� sortIndex_pred �֐��̕Ԃ�l��z��
' s_index �� vec �͈̔͊O�̒l�������͏d�����������ꍇ�̓���͖���`
' subV/subM �֐����g����葬���͂�
Sub permutate(ByRef vec As Variant, ByRef s_index As Variant)
    If Dimension(vec) <> 1 Or Dimension(s_index) <> 1 Then Exit Sub
    If rowSize(vec) = 0 Or rowSize(vec) <> rowSize(s_index) Then Exit Sub
    Dim i As Long, k As Long:    k = LBound(vec)
    Dim tmp As Variant
    If VarType(vec) = VarType(Array()) Then
        ReDim tmp(LBound(vec) To UBound(vec))
        For i = LBound(s_index) To UBound(s_index) Step 1
            swapVariant tmp(k), vec(s_index(i))
            k = k + 1
        Next i
        If swapVariant(tmp, vec) = 0 Then
            For i = LBound(vec) To UBound(vec) Step 1
                swapVariant tmp(i), vec(i)
            Next i
        End If
    ElseIf IsObject(vec(LBound(vec))) Then
        tmp = vec
        For i = LBound(s_index) To UBound(s_index) Step 1
            Set vec(k) = Nothing
            Set vec(k) = tmp(s_index(i))
            Set tmp(s_index(i)) = Nothing
            k = k + 1
        Next i
    Else
        tmp = vec
        For i = LBound(s_index) To UBound(s_index) Step 1
            vec(k) = tmp(s_index(i))
            k = k + 1
        Next i
    End If
End Sub

' permutate �ŕ��׊�����ꂽ1�����z������̏���ɖ߂�
Sub permutate_back(ByRef vec As Variant, ByRef s_index As Variant)
    Dim k As Long:          k = LBound(vec)
    Dim index2() As Long:   ReDim index2(LBound(vec) To UBound(vec))
    Dim z As Variant
    For Each z In s_index
        index2(z) = k
        k = k + 1
    Next z
    permutate vec, index2
End Sub

'�\�[�g�ςݔz�񂩂�w�肳�ꂽ�v�f�ȏ�̒l�������ŏ��̈ʒu���擾
Function lower_bound(ByRef matrix As Variant, ByRef val As Variant) As Variant
    lower_bound = lower_bound_imple(matrix, val, p_less, LBound(matrix, 1), 1 + UBound(matrix, 1))
End Function
    Public Function p_lower_bound(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_lower_bound = make_funPointer(AddressOf lower_bound, firstParam, secondParam)
    End Function

'�\�[�g�ςݔz�񂩂�w�肳�ꂽ�v�f�ȏ�̒l�������ŏ��̈ʒu���擾�i��r�֐��g�p�j
Function lower_bound_pred(ByRef matrix As Variant, ByRef val As Variant, ByRef pred As Variant) As Variant
    lower_bound_pred = lower_bound_imple(matrix, _
                                         val, _
                                         pred, _
                                         LBound(matrix, 1), 1 + UBound(matrix, 1))
End Function
    ' ����͒ʏ�̊֐��I�u�W�F�N�g�Ƃ͈قȂ�i��r�֐��݂̂������Ɏ��j
    ' mapF_swap(p_lower_bound_pred(comp), matrix, values) �Ƃ����g�p���@��z��
    Function p_lower_bound_pred(ByRef comp As Variant) As Variant
        p_lower_bound_pred = make_funPointer( _
                    AddressOf lower_bound_pred_zzz, _
                    Empty, _
                    make_funPointer(AddressOf makePair, yield_0, comp, 2))
    End Function
    Private Function lower_bound_pred_zzz(ByRef matrix As Variant, ByRef val_comp As Variant) As Variant
        lower_bound_pred_zzz = lower_bound_pred(matrix, val_comp(0), val_comp(1))
    End Function

'�\�[�g�ςݔz�񂩂�w�肳�ꂽ�v�f���傫���l�������ŏ��̈ʒu���擾
Function upper_bound(ByRef matrix As Variant, ByRef val As Variant) As Variant
    upper_bound = upper_bound_imple(matrix, val, p_less, LBound(matrix, 1), 1 + UBound(matrix, 1))
End Function
    Public Function p_upper_bound(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_upper_bound = make_funPointer(AddressOf upper_bound, firstParam, secondParam)
    End Function

'�\�[�g�ςݔz�񂩂�w�肳�ꂽ�v�f���傫���l�������ŏ��̈ʒu���擾�i��r�֐��g�p�j
Function upper_bound_pred(ByRef matrix As Variant, ByRef val As Variant, ByRef pred As Variant) As Variant
    upper_bound_pred = upper_bound_imple(matrix, _
                                         val, _
                                         pred, _
                                         LBound(matrix, 1), 1 + UBound(matrix, 1))
End Function
    ' ����͒ʏ�̊֐��I�u�W�F�N�g�Ƃ͈قȂ�i��r�֐��݂̂������Ɏ��j
    ' mapF_swap(p_upper_bound_pred(comp), matrix, values) �Ƃ����g�p���@��z��
    Function p_upper_bound_pred(ByRef comp As Variant) As Variant
        p_upper_bound_pred = make_funPointer( _
                    AddressOf upper_bound_pred_zzz, _
                    Empty, _
                    make_funPointer(AddressOf makePair, yield_0, comp, 2))
    End Function
    Private Function upper_bound_pred_zzz(ByRef matrix As Variant, ByRef val_comp As Variant) As Variant
        upper_bound_pred_zzz = upper_bound_pred(matrix, val_comp(0), val_comp(1))
    End Function

'lower_bound��upper_bound�̑g
Function equal_range(ByRef matrix As Variant, ByRef val As Variant) As Variant
    equal_range = VBA.Array(lower_bound(matrix, val), upper_bound(matrix, val))
End Function
    Public Function p_equal_range(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal_range = make_funPointer(AddressOf equal_range, firstParam, secondParam)
    End Function

'lower_bound��upper_bound�̑g�i��r�֐��g�p�j
Function equal_range_pred(ByRef matrix As Variant, ByRef val As Variant, ByRef pred As Variant) As Variant
    equal_range_pred = VBA.Array(lower_bound_pred(matrix, val, pred), upper_bound_pred(matrix, val, pred))
End Function
    ' ����͒ʏ�̊֐��I�u�W�F�N�g�Ƃ͈قȂ�i��r�֐��݂̂������Ɏ��j
    ' mapF_swap(p_equal_range_pred(comp), matrix, values) �Ƃ����g�p���@��z��
    Function p_equal_range_pred(ByRef comp As Variant) As Variant
        p_equal_range_pred = make_funPointer( _
                    AddressOf equal_range_pred_zzz, _
                    Empty, _
                    make_funPointer(AddressOf makePair, yield_0, comp, 2))
    End Function
    Private Function equal_range_pred_zzz(ByRef matrix As Variant, ByRef val_comp As Variant) As Variant
        equal_range_pred_zzz = equal_range_pred(matrix, val_comp(0), val_comp(1))
    End Function

'�\�[�g�ςݔz�񂩂�����ɂ���ċ敪������Ă���ʒu�̈ꗗ�𓾂�
Function partition_points(ByRef vec As Variant) As Variant
    partition_points = partition_points_pred(vec, p_less)
End Function

'�\�[�g�ςݔz�񂩂�����ɂ���ċ敪������Ă���ʒu�̈ꗗ�𓾂�
Function partition_points_pred(ByRef vec As Variant, ByRef pred As Variant) As Variant
    Dim ret As Variant
    ret = makeM(sizeof(vec) + 1)
    Dim rPos As Long:   rPos = LBound(vec)
    Dim wPos As Long:   wPos = 0
    Do
        ret(wPos) = rPos
        If UBound(vec) < rPos Then Exit Do
        rPos = upper_bound_imple(vec, vec(rPos), pred, rPos, 1 + UBound(vec))
        wPos = wPos + 1
    Loop
    ReDim Preserve ret(0 To wPos)
    swapVariant partition_points_pred, ret
End Function

'�q�� ������less
Function less_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    less_dic = 0&
    Dim i As Long
    For i = 0 To min_fun(sizeof(a), sizeof(b)) - 1 Step 1
        If getNth_b(a, i) < getNth_b(b, i) Then
            less_dic = 1&
            Exit For
        ElseIf getNth_b(a, i) > getNth_b(b, i) Then
            Exit For
        End If
    Next i
End Function
    Function p_less_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_less_dic = make_funPointer(AddressOf less_dic, firstParam, secondParam)
    End Function

'�q�� ������less_equal
Function less_equal_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    less_equal_dic = IIf(0 = less_dic(b, a), 1&, 0&)
End Function
    Function p_less_equal_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_less_equal_dic = make_funPointer(AddressOf less_equal_dic, firstParam, secondParam)
    End Function

'�q�� ������greater
Function greater_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    greater_dic = less_dic(b, a)
End Function
    Function p_greater_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_greater_dic = make_funPointer(AddressOf greater_dic, firstParam, secondParam)
    End Function

'�q�� ������greater_equal
Function greater_equal_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    greater_equal_dic = IIf(0 = less_dic(a, b), 1&, 0&)
End Function
    Function p_greater_equal_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_greater_equal_dic = make_funPointer(AddressOf greater_equal_dic, firstParam, secondParam)
    End Function

'�q�� ������equal
Function equal_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    equal_dic = IIf(0 = less_dic(a, b) And 0 = less_dic(b, a), 1&, 0&)
End Function
    Function p_equal_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal_dic = make_funPointer(AddressOf equal_dic, firstParam, secondParam)
    End Function

'�q�� ������notEqual
Function notEqual_dic(ByRef a As Variant, ByRef b As Variant) As Variant
    notEqual_dic = IIf(0 = equal_dic(b, a), 1&, 0&)
End Function
    Function p_notEqual_dic(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_notEqual_dic = make_funPointer(AddressOf notEqual_dic, firstParam, secondParam)
    End Function

'#####################
Private Function lower_bound_imple(ByRef matrix As Variant, _
                                   ByRef val As Variant, _
                                   ByRef comp As Variant, _
                                   ByVal begin_ As Long, _
                                   ByVal end_ As Long) As Long
    Dim mid_ As Long
    If end_ - begin_ < 8 Then
        Do While unbind_invoke(comp, matrix(begin_), val) 'And begin_ < end_
            begin_ = begin_ + 1
            If end_ <= begin_ Then Exit Do
        Loop
        lower_bound_imple = begin_
    Else
        mid_ = begin_ + CLng((end_ - begin_) / 2)
        If unbind_invoke(comp, matrix(mid_), val) Then
            lower_bound_imple = lower_bound_imple(matrix, val, comp, mid_, end_)
        Else
            lower_bound_imple = lower_bound_imple(matrix, val, comp, begin_, mid_)
        End If
    End If
End Function

Private Function upper_bound_imple(ByRef matrix As Variant, _
                                   ByRef val As Variant, _
                                   ByRef comp As Variant, _
                                   ByVal begin_ As Long, _
                                   ByVal end_ As Long) As Long
    Dim mid_ As Long
    If end_ - begin_ < 8 Then
        Do While 0 = unbind_invoke(comp, val, matrix(begin_)) 'And begin_ < end_
            begin_ = begin_ + 1
            If end_ <= begin_ Then Exit Do
        Loop
        upper_bound_imple = begin_
    Else
        mid_ = begin_ + CLng((end_ - begin_) / 2)
        If unbind_invoke(comp, val, matrix(mid_)) Then
            upper_bound_imple = upper_bound_imple(matrix, val, comp, begin_, mid_)
        Else
            upper_bound_imple = upper_bound_imple(matrix, val, comp, mid_, end_)
        End If
    End If
End Function
'####################

