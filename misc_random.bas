Attribute VB_Name = "misc_random"
'misc_random
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'********************************************************************
'   �m�����z
' Function  seed_Engine         �����V�[�h�ݒ�
' Function  uniform_int_dist    ��l����(Long)    (��, from, to)
' Function  uniform_real_dist   ��l����(Double)  (��, from, to)
' Function  normal_dist         ���K���z          (��, ����, �W���΍�)
' Function  bernoulli_dist      Bernoulli���z     (��, �����m��)
' Function  discrete_dist       ���U���z          (��, �����䗦�z��)
' Function  random_iota         iota�̃����_����
' Function  random_shuffle      �z��̗v�f�������_���ɕ��ёւ����z����o��
'********************************************************************

Declare PtrSafe Function seed_Engine Lib "mapM.dll" _
                                    (Optional ByRef seedN As Variant) As Long

Declare PtrSafe Function uniform_int_dist Lib "mapM.dll" _
                                    (ByVal N As Long, _
                                ByVal fromN As Long, _
                            ByVal toN As Long) As Variant

Declare PtrSafe Function uniform_real_dist Lib "mapM.dll" _
                                        (ByVal N As Long, _
                                    ByVal fromD As Double, _
                                ByVal toD As Double) As Variant

Declare PtrSafe Function normal_dist Lib "mapM.dll" _
                                    (ByVal N As Long, _
                                ByVal mean As Double, _
                            ByVal stddev As Double) As Variant
            
Declare PtrSafe Function bernoulli_dist Lib "mapM.dll" _
                                    (ByVal N As Long, _
                                ByVal prob As Double) As Variant

' ��l����(Long) : (��N, [�͈�])
    Private Function uniform_int_dist_(ByRef N As Variant, ByRef fromto As Variant) As Variant
        uniform_int_dist_ = uniform_int_dist(N, fromto(LBound(fromto)), fromto(UBound(fromto)))
    End Function
Function p_uniform_int_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_uniform_int_dist = make_funPointer(AddressOf uniform_int_dist_, firstParam, secondParam)
End Function

' ��l����(Double) : (��N, [�͈�])
    Private Function uniform_real_dist_(ByRef N As Variant, ByRef fromto As Variant) As Variant
        uniform_real_dist_ = uniform_real_dist(N, fromto(LBound(fromto)), fromto(UBound(fromto)))
    End Function
Function p_uniform_real_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_uniform_real_dist = make_funPointer(AddressOf uniform_real_dist_, firstParam, secondParam)
End Function

' ���K���z : (��N, [����,�W���΍�])
    Private Function normal_dist_(ByRef N As Variant, ByRef meandev As Variant) As Variant
        normal_dist_ = normal_dist(N, meandev(LBound(meandev)), meandev(UBound(meandev)))
    End Function
Function p_normal_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_normal_dist = make_funPointer(AddressOf normal_dist_, firstParam, secondParam)
End Function

' Bernoulli���z : (��N, �����m��)
    Private Function bernoulli_dist_(ByRef N As Variant, ByRef prob As Variant) As Variant
        bernoulli_dist_ = bernoulli_dist(N, prob)
    End Function
Function p_bernoulli_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_bernoulli_dist = make_funPointer(AddressOf bernoulli_dist_, firstParam, secondParam)
End Function

' ���U���z : (�� N, �����䗦�z�� Pi)
' �����䗦�͔񕉂̎����A���v�� 1 �ɂȂ�Ȃ��Ă���
' �Ԃ�l�͒��� N �̔z��ŁA�e�v�f�� 0�`sizeof(�����䗦�z��)-1 �̐���
' ����i�̔����䗦 �` Pi �ƂȂ镪�z�i������ LBound(Pi) = 0 �Ɖ���j
Function discrete_dist(ByRef N As Variant, ByRef probs As Variant) As Variant
    Dim segments As Variant, distribution As Variant
    segments = scanl1(p_plus, probs)
    distribution = uniform_real_dist(N, 0#, segments(UBound(segments)))
    If 0 < N Then
        discrete_dist = foldl1(p_plus, product_set(p_less, segments, distribution), 1)
    Else
        discrete_dist = foldl1(p_plus, mapF(p_less(, distribution), segments))
    End If
End Function
    Function p_discrete_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_discrete_dist = make_funPointer(AddressOf discrete_dist, firstParam, secondParam)
    End Function

' iota�̃����_���Łifrom_i����to_i�܂ł̎��R���������_���ɕ��ׂ��x�N�g���j
' Fisher-Yates
Function random_iota(ByVal from_i As Long, ByVal to_i As Long) As Variant
    Dim ret As Variant, i As Long, j As Long, tmp As Long
    ret = iota(from_i, to_i)
    For i = UBound(ret) To 1 Step -1
        j = uniform_int_dist(0, 0, i)
        tmp = ret(i): ret(i) = ret(j): ret(j) = tmp
    Next i
    Call swapVariant(random_iota, ret)
End Function

' �z��̗v�f�������_���ɕ��ёւ����z����o��
Function random_shuffle(ByRef vec As Variant, Optional ByRef dummy As Variant) As Variant
    random_shuffle = vec
    Call permutate(random_shuffle, random_iota(LBound(vec), UBound(vec)))
End Function
    Function p_random_shuffle(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_random_shuffle = make_funPointer(AddressOf random_shuffle, firstParam, secondParam)
    End Function
