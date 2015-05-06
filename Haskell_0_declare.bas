Attribute VB_Name = "Haskell_0_declare"
'Haskell_0_declare
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'======================================================
'          API宣言
'   Function Dimension          配列の次元取得
'   Function placeholder        プレースホルダ・オブジェクトの生成
'   Function is_placeholder     プレースホルダ・オブジェクト判定
'   Function unbind_invoke      bindされていないVBA関数を2引数で呼び出す
'   Function mapF_imple         配列matrixの各要素elemにCallback関数を適用する
'   Function zipWith            2つの配列の各要素に関数を適用する
'   Function foldl              配列に対する特定の軸に沿った左畳み込み（初期値指定あり）
'   Function foldr              配列に対する特定の軸に沿った右畳み込み（初期値指定あり）
'   Function foldl1             配列に対する特定の軸に沿った左畳み込み（先頭要素を初期値とする）
'   Function foldr1             配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
'   Function scanl              配列に対する特定の軸に沿った左scan（初期値指定あり）
'   Function scanr              配列に対する特定の軸に沿った右scan（初期値指定あり）
'   Function scanl1             配列に対する特定の軸に沿った左scan（先頭要素を初期値とする）
'   Function scanr1             配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
'   Function stdsort            1次元配列のソートインデックス出力
'   Function find_imple         述語による検索
'   Function repeat_imple       関数適用のループ（+ 終了条件）
'   Function moveVariant        sourceのVARIANT変数を返り値のVARIANTへmoveする
'======================================================
' Callbackとして使える関数のシグネチャは
' Function fun(ByRef x As Variant, ByRef y As Variant) As Variant
' もしくは
' Function fun(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
'======================================================
' VBA配列の次元取得
Declare Function Dimension Lib "mapM.dll" (ByRef v As Variant) As Long

'プレースホルダ・オブジェクトの生成
Declare Function placeholder Lib "mapM.dll" () As Variant

'プレースホルダ・オブジェクト判定
Declare Function is_placeholder Lib "mapM.dll" (ByRef v As Variant) As Long

'bindされていないVBA関数を2引数で呼び出す
Declare Function unbind_invoke Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef param1 As Variant, _
        ByRef param2 As Variant) As Variant

' 配列matrixの各要素elemにCallback関数を適用する
Declare Function mapF_imple Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant) As Variant

'配列matrix1とmatrix2の各要素に2変数のCallback関数を適用する
Declare Function zipWith Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix1 As Variant, _
        ByRef matrix2 As Variant) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った左畳み込み（初期値指定あり）
Declare Function foldl Lib "mapM.dll" ( _
                    ByRef pCallback As Variant, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った右畳み込み（初期値指定あり）
Declare Function foldr Lib "mapM.dll" ( _
                    ByRef pCallback As Variant, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った左畳み込み（先頭要素を初期値とする）
Declare Function foldl1 Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
Declare Function foldr1 Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った左scan（初期値指定あり）
Declare Function scanl Lib "mapM.dll" ( _
                    ByRef pCallback As Variant, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った右scan（初期値指定あり）
Declare Function scanr Lib "mapM.dll" ( _
                    ByRef pCallback As Variant, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った左scan（先頭要素を初期値とする）
Declare Function scanl1 Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
Declare Function scanr1 Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 1次元配列のソートインデックス出力
Declare Function stdsort Lib "mapM.dll" (ByRef ary As Variant, _
                                         ByVal defaultFlag As Long, _
                                         ByRef pComp As Variant) As Variant

' 述語による検索
Declare Function find_imple Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        ByVal def As Long) As Long

'関数適用のループ（+ 終了条件）
Declare Function repeat_imple Lib "mapM.dll" ( _
                        ByRef init As Variant, _
                    ByRef pred As Variant, _
                ByRef trans As Variant, _
            ByVal maxN As Long, _
        ByVal scan As Long, _
    ByVal stopCondition As Long) As Variant

'sourceのVARIANT変数を返り値のVARIANTへmoveする
Declare Function moveVariant Lib "mapM.dll" (ByRef source As Variant) As Variant
