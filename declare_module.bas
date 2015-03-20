Attribute VB_Name = "declare_module"
'declare_module
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'======================================================
'          API宣言
'   Function Dimension          配列の次元取得
'   Function placeholder        プレースホルダ・オブジェクトの生成
'   Function is_placeholder     プレースホルダ・オブジェクト判定
'   Function simple_invoke      2引数に関数を適用する
'   Function mapL               配列の各要素と引数に関数を適用する（ユーザコードでは主にmapFを使用する）
'   Function mapR               引数と配列の各要素に関数を適用する（ユーザコードでは主にmapFを使用する）
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

' 2引数にCallback関数を適用する
Declare Function simple_invoke Lib "mapM.dll" ( _
                ByVal pCallback As Long, _
            Optional ByRef param1 As Variant, _
        Optional ByRef param12 As Variant) As Variant

' 配列matrixの各要素elemとparamにCallback(elem, param)を適用する（ユーザコードでは主にmapFを使用する）
Declare Function mapL Lib "mapM.dll" ( _
                ByVal pCallback As Long, _
            ByRef matrix As Variant, _
        Optional ByRef param As Variant) As Variant

' paramと配列matrixの各要素elemにCallback(param, elem)を適用する（ユーザコードでは主にmapFを使用する）
Declare Function mapR Lib "mapM.dll" ( _
                ByVal pCallback As Long, _
            ByRef param As Variant, _
        ByRef matrix As Variant) As Variant

'配列matrix1とmatrix2の各要素に2変数のCallback関数を適用する
Declare Function zipWith Lib "mapM.dll" ( _
                ByVal pCallback As Long, _
            ByRef matrix1 As Variant, _
        ByRef matrix2 As Variant) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った左畳み込み（初期値指定あり）
Declare Function foldl Lib "mapM.dll" ( _
                    ByVal pCallback As Long, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った右畳み込み（初期値指定あり）
Declare Function foldr Lib "mapM.dll" ( _
                    ByVal pCallback As Long, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った左畳み込み（先頭要素を初期値とする）
Declare Function foldl1 Lib "mapM.dll" ( _
                ByVal pCallback As Long, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
Declare Function foldr1 Lib "mapM.dll" ( _
                ByVal pCallback As Long, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った左scan（初期値指定あり）
Declare Function scanl Lib "mapM.dll" ( _
                    ByVal pCallback As Long, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った右scan（初期値指定あり）
Declare Function scanr Lib "mapM.dll" ( _
                    ByVal pCallback As Long, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った左scan（先頭要素を初期値とする）
Declare Function scanl1 Lib "mapM.dll" ( _
                ByVal pCallback As Long, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3次元までのVBA配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
Declare Function scanr1 Lib "mapM.dll" ( _
                ByVal pCallback As Long, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 1次元配列のソートインデックス出力
Declare Function stdsort Lib "mapM.dll" (ByRef ary As Variant, ByVal pComp As Long) As Variant
