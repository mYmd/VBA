VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vh_pipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'vh_pipe
'Copyright (c) 2016 mmYYmmdd

'*********************************************************************************
'   パイプライン演算子 |>
'   このオブジェクトは直接Newするのではなく
'   ファクトリ関数misc_utility.pipeまたはmisc_utility.pipe_を使用して生成する
' --------------------------------------------------------------------------------
'   Function        x                   関数適用して自分自身を返す
'   Function        →                  関数適用して自分自身を返す
'   Property Get    val                 結果を取得
'   Property Get    pop                 結果を取得(値はmoveされる）
'   Function        swap                vh_pipeどうしのswap
'   Function        clone               同じ値を持ったオブジェクトの生成
'*********************************************************************************
    Private ret As Variant

    Friend Sub swap_val_(ByRef x As Variant)
        swapVariant ret, x
    End Sub

Public Function x(ByRef fun As Variant) As vh_pipe
    swapVariant ret, applyFun(ret, fun)
    Set x = Me
End Function

Public Function →(ByRef fun As Variant) As vh_pipe
    swapVariant ret, applyFun(ret, fun)
    Set → = Me
End Function

Public Property Get val() As Variant
    val = ret
End Property

Public Property Get pop() As Variant
    swapVariant pop, ret
End Property

Public Function swap(ByVal other As vh_pipe) As vh_pipe
    other.swap_val_ ret
    Set swap = Me
End Function

Public Function clone() As vh_pipe
    Set clone = New vh_pipe
    clone.swap_val_ (ret)    'コピーを渡す
End Function
