VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vh_stdvec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'vh_stdvec
'Copyright (c) 2016 mmYYmmdd

'*********************************************************************************
'   標準ベクトル
' --------------------------------------------------------------------------------
'   Function        create               1次元配列をmoveしてVECTORを作成
'   Function        clear()                 全ての要素を削除
'   Function        free                配列部分を戻して自身は解放
'   Property Get    size()              書き込み済み要素数を取得
'   Property Get    capacity            ReDimせずに格納できる最大の要素数を取得
'   Function        shrink              領域を size まで切り詰める
'   Sub             printS              デバッグウィンドウに size と capacity を表示する
'   Sub             printM              デバッグウィンドウに 書き込み済み領域全体を表示する
'   Sub             printM_             デバッグウィンドウにジャグ配列を展開して表示する
'   Property Get    val                 指定位置の値を取得
'   Property Get    pop                 valのmove版
'   Function        push                指定位置の値を設定
'   Function        push_n              同一の値に対するpushの複数回適用
'   Function        push_array          配列の一括 push
'   Function        swap                pushのswap版
'   Function        swap_array          push_arrayのswap版
'   Function        move                pushのmove版
'   Function        move_array          push_arrayのmove版
'   Function        insert_empty        beginの位置の前にn個のEmptyを挿入
'   Function        push_back           書き込み済み末尾に対するpush（領域拡張する）
'   Function        push_back_n         書き込み済み末尾に対する同一値の複数push（領域拡張する）
'   Function        push_back_array     書き込み済み末尾に対するpush_array（領域拡張する）
'   Function        move_back           書き込み済み末尾に対するmove（領域拡張する）
'   Function        move_back_array     書き込み済み末尾に対するmove_array（領域拡張する）
'*********************************************************************************

Private vec_ As Variant
Private pos_write As Long

' 1次元配列をmoveしてVECTORを作成（LBound = 0にする）
Public Function create(ByRef v As Variant) As Variant
    If Dimension(v) = 1 Then
        vec_ = Empty
        swapVariant vec_, v
        changeLBound vec_, 0
        pos_write = UBound(vec_)
    Else
        clear
    End If
    Set create = Me
End Function

' 全ての要素を削除
Public Function clear() As vh_stdvec
    vec_ = VBA.Array()
    pos_write = -1
    Set clear = Me
End Function

' 配列部分を戻して自身は解放
Public Function free() As Variant
    init_
    shrink
    swapVariant free, vec_
    init_
End Function

' 書き込み済み要素数を取得
Public Property Get size() As Long
    init_
    size = 1 + pos_write
End Property

' ReDimせずに格納できる最大の要素数を取得
Public Property Get capacity() As Long
    init_
    capacity = rowSize(vec_)
End Property

' 領域を size まで切り詰める
Public Function shrink() As vh_stdvec
    init_
    If 0 <= pos_write Then ReDim Preserve vec_(0 To pos_write)
    Set shrink = Me
End Function

' デバッグウィンドウに size と capacity を表示する
Public Sub printS()
    init_
    Debug.Print "[Data]: 0 -> " & pos_write & " : [Capacity]: 0 -> " & UBound(vec_)
End Sub

' デバッグウィンドウに 書き込み済み領域全体を表示する
Public Sub printM(Optional ByRef r As Variant)
    init_
    If 0 < rowSize(vec_) Then
        Dim n As Long
        If IsMissing(r) Then
            n = pos_write
        Else
            n = min_fun(r, pos_write)
        End If
        Haskell_3_printM.printM subV(vec_, a__a(0, n))
    Else
        Debug.Print "#Empty Vector#"
    End If
End Sub

' デバッグウィンドウにジャグ配列を展開して表示する
Public Sub printM_(Optional ByVal r As Variant)
    init_
    If 0 < rowSize(vec_) Then
        Dim n As Long
        If IsMissing(r) Then
            n = pos_write
        Else
            n = min_fun(r, pos_write)
        End If
        Haskell_3_printM.printM_ subV(vec_, a__a(0, n))
    Else
        Debug.Print "#Empty Vector#"
    End If
End Sub

' 指定位置の 値 を取得する（1以上の第2引数を指定した場合はその長さの配列）
Public Property Get val(ByVal begin As Long, Optional ByVal len_ As Long = 0) As Variant
    If 0 < len_ Then
        Dim ret As Variant, i As Long
        ret = makeM(len_)
        For i = 0 To len_ - 1 Step 1
            ret(i) = vec_(begin + i)
        Next i
        swapVariant val, ret
    Else
        val = vec_(begin)
    End If
End Property

' 指定位置の 値 を取得する move版
Public Property Get pop(ByVal begin As Long, Optional ByVal len_ As Long = 0) As Variant
    If 0 < len_ Then
        Dim ret As Variant, i As Long
        ret = makeM(len_)
        For i = 0 To len_ - 1 Step 1
            swapVariant ret(i), vec_(begin + i)
        Next i
        swapVariant pop, ret
    Else
        swapVariant pop, vec_(begin)
    End If
End Property

' 指定位置の 値 を設定する
Public Function push(ByVal pos As Long, ByVal x As Variant) As vh_stdvec
    Set push = Me.swap(pos, x)
End Function

' 同一の値に対する push の複数回適用
Public Function push_n(ByVal pos As Long, ByRef x As Variant, ByVal n As Long) As vh_stdvec
    If 0 < n Then
        Set push_n = Me.swap_array(pos, repeat(x, n))
    Else
        Set push_n = Me
    End If
End Function

' 配列の一括 push
Public Function push_array(ByVal pos As Long, ByVal x As Variant) As vh_stdvec
    Set push_array = swap_array(pos, x)
End Function

' pushのswap版
Public Function swap(ByVal pos As Long, ByRef x As Variant) As vh_stdvec
    swapVariant vec_(pos), x
    If pos_write < pos Then pos_write = pos
    Set swap = Me
End Function

' push_arrayのswap版
Public Function swap_array(ByVal pos As Long, ByRef x As Variant) As vh_stdvec
    Dim i As Long
    Dim kkk As Long: kkk = UBound(x)
    For i = pos + sizeof(x) - 1 To pos Step -1
        Call Me.swap(i, x(kkk))
        kkk = kkk - 1
    Next i
    Set swap_array = Me
End Function

' pushのmove版
Public Function move(ByVal pos As Long, ByRef x As Variant) As vh_stdvec
    vec_(pos) = Empty
    Set move = Me.swap(pos, x)
End Function

' push_arrayのmove版
Public Function move_array(ByVal pos As Long, ByRef x As Variant) As vh_stdvec
    Dim i As Long
    For i = pos + sizeof(x) - 1 To pos Step -1
        vec_(i) = Empty
    Next i
    Set move_array = swap_array(pos, x)
End Function

' beginの位置の前にn個のEmptyを挿入する
Public Function insert_empty(ByVal begin As Long, ByVal n As Long) As vh_stdvec
    init_
    If begin < 0 Or n < 1 Then
        Set insert_empty = Me
        Exit Function
    End If
    Dim NNN As Long
    NNN = max_fun(begin, pos_write) + n - 1
    If UBound(vec_) < NNN Then
        ReDim Preserve vec_(0 To max_fun(1 + 2 * UBound(vec_), NNN))
    End If
    If begin <= pos_write Then
        Dim b As Long, e As Long
        b = pos_write
        e = b + n
        If VarType(vec_) = VarType(Array()) Then
            Do While begin <= b
                swapVariant vec_(e), vec_(b)
                vec_(b) = Empty
                e = e - 1
                b = b - 1
            Loop
        ElseIf IsObject(vec_(0)) Then
            Do While begin <= b
                Set vec_(e) = vec_(b)
                Set vec_(b) = Nothing
                e = e - 1
                b = b - 1
            Loop
        Else
            Do While begin <= b
                vec_(e) = vec_(b)
                vec_(b) = Empty
                e = e - 1
                b = b - 1
            Loop
        End If
        pos_write = pos_write + n
    End If
    Set insert_empty = Me
End Function

' 書き込み済み末尾に対するpush（領域拡張する）
Public Function push_back(ByVal x As Variant) As vh_stdvec
    Set push_back = move_back(x)
End Function

' 書き込み済み末尾に対する同一値の複数push（領域拡張する）
Public Function push_back_n(ByRef x As Variant, ByVal n As Long) As vh_stdvec
    If 0 < n Then
        Set push_back_n = move_back_array(repeat(x, n))
    Else
        Set push_back_n = Me
    End If
End Function

' 書き込み済み末尾に対するpush_array（領域拡張する）
Public Function push_back_array(ByVal arr As Variant) As vh_stdvec
    Set push_back_array = move_back_array(arr)
End Function

' 書き込み済み末尾に対するmove（領域拡張する）
Public Function move_back(ByRef x As Variant) As vh_stdvec
    Set move_back = Me.insert_empty(pos_write + 1, 1).swap(pos_write + 1, x)
End Function

' 書き込み済み末尾に対するmove_array（領域拡張する）
Public Function move_back_array(ByRef arr As Variant) As vh_stdvec
    If Dimension(arr) = 1 And 0 < rowSize(arr) Then
        Set move_back_array = insert_empty(pos_write + 1, rowSize(arr)).swap_array(pos_write + 1, arr)
    Else
        Set move_back_array = Me
    End If
End Function

'-----------------------------------------------
    Private Sub init_()
        If IsEmpty(vec_) Then
            vec_ = VBA.Array()
            pos_write = -1
        End If
    End Sub
