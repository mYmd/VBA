Attribute VB_Name = "misc_random"
'misc_random
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'********************************************************************
'   確率分布
' Function  seed_Engine         乱数シード設定
' Function  uniform_int_dist    一様整数(Long)    (個数, from, to)
' Function  uniform_real_dist   一様実数(Double)  (個数, from, to)
' Function  normal_dist         正規分布          (個数, 平均, 標準偏差)
' Function  bernoulli_dist      Bernoulli分布     (個数, 発生確率)
' Function  discrete_dist       離散分布          (個数, 発生比率配列)
' Function  random_shaffle      配列の要素をランダムに並び替える。
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

' 一様整数(Long) : (個数N, [範囲])
    Private Function uniform_int_dist_(ByRef N As Variant, ByRef fromto As Variant) As Variant
        uniform_int_dist_ = uniform_int_dist(N, fromto(LBound(fromto)), fromto(UBound(fromto)))
    End Function
Function p_uniform_int_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_uniform_int_dist = make_funPointer(AddressOf uniform_int_dist_, firstParam, secondParam)
End Function

' 一様実数(Double) : (個数N, [範囲])
    Private Function uniform_real_dist_(ByRef N As Variant, ByRef fromto As Variant) As Variant
        uniform_real_dist_ = uniform_real_dist(N, fromto(LBound(fromto)), fromto(UBound(fromto)))
    End Function
Function p_uniform_real_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_uniform_real_dist = make_funPointer(AddressOf uniform_real_dist_, firstParam, secondParam)
End Function

' 正規分布 : (個数N, [平均,標準偏差])
    Private Function normal_dist_(ByRef N As Variant, ByRef meandev As Variant) As Variant
        normal_dist_ = normal_dist(N, fromto(LBound(meandev)), fromto(UBound(meandev)))
    End Function
Function p_normal_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_normal_dist = make_funPointer(AddressOf normal_dist_, firstParam, secondParam)
End Function

' Bernoulli分布 : (個数N, 発生確率)
    Private Function bernoulli_dist_(ByRef N As Variant, ByRef prob As Variant) As Variant
        bernoulli_dist_ = bernoulli_dist(N, prob)
    End Function
Function p_bernoulli_dist(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_bernoulli_dist = make_funPointer(AddressOf bernoulli_dist_, firstParam, secondParam)
End Function

' 離散分布 : (個数 N, 発生比率配列 Pi)
' 発生比率は非負の実数、合計が 1 にならなくても可
' 返り値は長さ N の配列で、各要素は 0～sizeof(発生比率配列)-1 の整数
' 整数iの発生比率 ～ Pi となる分布（ただし LBound(Pi) = 0 と仮定）
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

' 配列の要素をランダムに並び替える。
Function random_shaffle(ByRef vec As Variant, Optional ByRef dummy As Variant) As Variant
    Dim index As Variant
    index = sortIndex(uniform_real_dist(sizeof(vec), 0, 1))
    If 0 <> LBound(vec) Then index = mapF(p_plus(LBound(vec)), index)
    random_shaffle = vec
    permutate random_shaffle, index
End Function
    Function p_random_shaffle(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_random_shaffle = make_funPointer(AddressOf random_shaffle, firstParam, secondParam)
    End Function
    
