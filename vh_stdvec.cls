VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vh_stdvec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'vh_stdvec
'Copyright (c) 2016 mmYYmmdd

'*********************************************************************************
'   標準ベクトル
' --------------------------------------------------------------------------------
'   Function        from                Variantの1次元配列をmoveしてVECTORを作成
'   Function        clear               全ての要素を削除
'   Function        free                配列部分を戻して自身は解放
'   Function        swap                vh_stdvecどうしのswap
'   Function        clone               同じ配列を持ったオブジェクトの生成
'   Property Get    size                書き込み済み要素数を取得
'   Property Get    capacity            ReDimせずに格納できる最大の要素数を取得
'   Function        shrink              領域を指定の長さまたはsizeまで切り詰める
'   Function        erase_n             指定した位置から個数を指定して範囲を削除する
'   Sub             printS              デバッグウィンドウに size と capacity を表示する
'   Sub             printM              デバッグウィンドウに 書き込み済み領域全体を表示する
'   Sub             printM_             デバッグウィンドウにジャグ配列を展開して表示する
'   Property Get    val                 指定位置の値を取得
'   Property Get    pop                 valのmove版（値はmoveされる）
'   Function        push                指定位置に値を設定（値はmoveされる）
'   Function        push_n              同一の値に対するpushの複数回適用
'   Function        push_array          配列の一括 push（配列の各要素はmoveされる）
'   Function        insert_empty        beginの位置の前にn個のEmptyを挿入（領域拡張する）
'   Function        push_back           書き込み済み末尾に対するpush（領域拡張する）
'   Function        push_back_n         書き込み済み末尾に対するpush_n（領域拡張する）
'   Function        push_back_array     書き込み済み末尾に対するpush_array（領域拡張する）
'   Function        push_back_arrays    push_back_arrayの反復
'   Function        filter              抽出条件を与えて配列をフィルタリングする
'   Function        map                 変換式を与えて配列を変換する
'   Function        orderby             比較関数を与えて配列をソートする
'   Function        zipWith             他のベクトルとの zipWith
'   Function        inner_join          他のベクトルとの Inner Join
'*********************************************************************************
    Private Declare PtrSafe Function copy_valueMember Lib "mapM.dll" _
                                            (ByVal source As vh_stdvec, _
                                        ByRef mem As Variant, _
                                    ByVal Target As vh_stdvec, _
                                ByVal direc As Long) As Boolean

    Private vec_ As Variant
    Private pos_write As Long

    Private Sub Class_Initialize()
        ReDim vec_(0 To 0)
        pos_write = -1
    End Sub
    
    Private Sub Class_Terminate()
        vec_ = Empty
    End Sub

' Variantの1次元配列をmoveして自身のデータとする（LBound = 0にする）
' 2次元配列を渡した場合は zipC する
Public Function from(ByRef v As Variant) As vh_stdvec
    If VarType(v) = VarType(Array()) Then
        vec_ = Empty
        Select Case Dimension(v)
        Case 1
            swapVariant vec_, v
            changeLBound vec_, 0
            pos_write = UBound(vec_)
        Case 2
            vec_ = zipC(v)
            changeLBound vec_, 0
            pos_write = UBound(vec_)
        Case Else
            Err.Raise Number:=575, Description:="vh_stdvec::from - inproper parmater"
        End Select
    Else
        Err.Raise Number:=575, Description:="vh_stdvec::from - inproper parmater"
    End If
    Set from = Me
End Function

' 全ての要素を削除
Public Function clear() As vh_stdvec
    ReDim vec_(0 To 0)
    pos_write = -1
    Set clear = Me
End Function

' 配列部分を戻して自身は解放
Public Function free() As Variant
    shrink
    swapVariant free, vec_
    ReDim vec_(0 To 0)
    pos_write = -1
End Function

'vh_stdvecどうしのswap
Public Function swap(ByVal Target As vh_stdvec) As vh_stdvec
    Call copy_valueMember(Me, vec_, Target, 0)
    Call copy_valueMember(Me, pos_write, Target, 0)
    Set swap = Me
End Function

' 同じ配列を持ったオブジェクトの生成
Public Function clone() As vh_stdvec
    Set clone = New vh_stdvec
    Call copy_valueMember(Me, vec_, clone, 1)
    Call copy_valueMember(Me, pos_write, clone, 1)
End Function

' 書き込み済み要素数を取得
Public Property Get size() As Long
    size = 1 + pos_write
End Property

' ReDimせずに格納できる最大の要素数を取得
Public Property Get capacity() As Long
    capacity = rowSize(vec_)
End Property

' 領域を指定の長さまたはsizeまで切り詰める
Public Function shrink(Optional ByVal size_ As Long = -1) As vh_stdvec
    If 0 < size_ Then pos_write = min_fun(pos_write, size_ - 1)
    If 0 <= pos_write Then ReDim Preserve vec_(0 To pos_write)
    Set shrink = Me
End Function

' 指定した位置から個数を指定して範囲を削除する
Public Function erase_n(ByVal begin As Long, ByVal n As Long) As vh_stdvec
    If 0 <= begin And begin <= pos_write And 0 < n Then
        Dim kkk As Long
        kkk = begin + n
        If begin < kkk Then
            Do While kkk <= pos_write
                swapVariant vec_(begin), vec_(kkk)
                begin = begin + 1
                kkk = kkk + 1
            Loop
        End If
        kkk = begin
        Do While kkk <= pos_write
            vec_(kkk) = Empty
            kkk = kkk + 1
        Loop
        pos_write = begin - 1
    End If
    Set erase_n = Me
End Function

' デバッグウィンドウに書き込み済み領域と配列領域を表示する
Public Sub printS()
    Debug.Print "[Data]: 0 -> " & pos_write & " : [Accessible]: 0 -> " & UBound(vec_)
End Sub

' デバッグウィンドウに 書き込み済み領域の全体または一部を表示する
Public Sub printM(Optional ByRef r As Variant)
    If 0 < rowSize(vec_) Then
        If IsMissing(r) Then
            Haskell_3_printM.printM vec_, VBA.Array(0, pos_write)
        ElseIf 0 < r Then
            Haskell_3_printM.printM vec_, VBA.Array(0, min_fun(r - 1, pos_write))
        Else
            Haskell_3_printM.printM vec_, VBA.Array(max_fun(pos_write + r + 1, 0), pos_write)
        End If
    Else
        Debug.Print "#Empty Vector#"
    End If
End Sub

' デバッグウィンドウにジャグ配列を展開して表示する
Sub printM_(Optional ByRef r As Variant, Optional ByRef c As Variant)
    Dim begin_ As Long, end_   As Long
    If IsMissing(r) Then
        begin_ = 0
        end_ = pos_write
    ElseIf 0 <= r Then
        begin_ = 0
        end_ = min_fun(pos_write, r - 1)
    Else
        end_ = pos_write
        begin_ = max_fun(0, end_ + r + 1)
    End If
    Haskell_3_printM.printM_ vec_, VBA.Array(begin_, end_), c
End Sub

' 指定位置の 値 を取得する（1以上の第2引数を指定した場合はその長さ分配列で返す）
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

' 指定位置の 値 を取得する move版（1以上の第2引数を指定した場合はその長さ分配列で返す）
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

' 指定位置に値を設定する（値はmoveされる）
Public Function push(ByVal pos As Long, ByRef x As Variant) As vh_stdvec
    vec_(pos) = Empty
    swapVariant vec_(pos), x
    If pos_write < pos Then pos_write = pos
    Set push = Me
End Function

' 同一の値に対する push の複数回適用
Public Function push_n(ByVal pos As Long, ByRef x As Variant, ByVal n As Long) As vh_stdvec
    If 0 < n Then
        Set push_n = Me.push_array(pos, repeat(x, n))
    Else
        Set push_n = Me
    End If
End Function

' 配列の一括 push（配列はmoveされる）
Public Function push_array(ByVal pos As Long, ByRef x As Variant) As vh_stdvec
    Dim i As Long
    If Dimension(x) = 1 Then
        Dim kkk As Long: kkk = UBound(x)
        For i = pos + sizeof(x) - 1 To pos Step -1
            Call Me.push(i, x(kkk))
            kkk = kkk - 1
        Next i
        x = Empty
    End If
    Set push_array = Me
End Function

' beginの位置の前にn個のEmptyを挿入する
Public Function insert_empty(ByVal begin As Long, ByVal n As Long) As vh_stdvec
    If begin < 0 Or n < 1 Then
        Set insert_empty = Me
        Exit Function
    End If
    Dim NNN As Long
    NNN = max_fun(begin, pos_write) + n
    If UBound(vec_) < NNN Then
        ReDim Preserve vec_(0 To max_fun(1 + 2 * UBound(vec_), NNN))
    End If
    If begin <= pos_write Then
        Dim b As Long, e As Long
        b = pos_write
        e = b + n
        Do While begin <= b
            swapVariant vec_(e), vec_(b)
            vec_(b) = Empty
            e = e - 1
            b = b - 1
        Loop
        pos_write = pos_write + n
    End If
    Set insert_empty = Me
End Function

' 書き込み済み末尾に対するpush（領域拡張する）
Public Function push_back(ByRef x As Variant) As vh_stdvec
    Set push_back = Me.insert_empty(pos_write + 1, 1).push(pos_write + 1, x)
End Function

' 書き込み済み末尾に対する同一値の複数push（領域拡張する）
Public Function push_back_n(ByRef x As Variant, ByVal n As Long) As vh_stdvec
    If 0 < n Then
        Set push_back_n = push_back_array(repeat(x, n))
    Else
        Set push_back_n = Me
    End If
End Function

' 書き込み済み末尾に対するpush_array（領域拡張する）
Public Function push_back_array(ByRef arr As Variant) As vh_stdvec
    If Dimension(arr) = 1 And 0 < rowSize(arr) Then
        Set push_back_array = insert_empty(pos_write + 1, rowSize(arr)).push_array(pos_write + 1, arr)
    Else
        Set push_back_array = Me
    End If
End Function

' push_back_arrayの反復
Public Function push_back_arrays(ByRef arrays As Variant) As vh_stdvec
    Dim i As Long
    For i = LBound(arrays) To UBound(arrays) Step 1
        push_back_array arrays(i)
    Next i
    Set push_back_arrays = Me
End Function

' 抽出条件を与えて配列をフィルタリングする
Public Function filter(ByRef pred As Variant) As vh_stdvec
    shrink
    Set filter = New vh_stdvec
    filter.from filterR(vec_, mapF(pred, vec_))
End Function

' 変換式を与えて配列を変換する
Public Function map(ByRef fun As Variant) As vh_stdvec
    shrink
    Set map = New vh_stdvec
    map.from mapF(fun, vec_)
End Function

' 比較関数を与えて配列をソートする
Public Function orderby(Optional ByRef comp As Variant) As vh_stdvec
    shrink
    Dim tmp As Variant
    tmp = vec_
    If IsMissing(comp) Then
        permutate tmp, sortIndex(tmp)       ' p_less
    Else
        permutate tmp, sortIndex_pred(tmp, comp)
    End If
    Set orderby = New vh_stdvec
    orderby.from tmp
End Function

' 他のベクトルとの zipWith
Public Function zipWith(ByRef fun As Variant, ByVal other As vh_stdvec) As vh_stdvec
    shrink
    Set zipWith = New vh_stdvec
    zipWith.from Haskell_0_declare.zipWith(fun, vec_, other.clone().free())
End Function

' 他のベクトルとの Inner Join
' key1 : 自身のキー列番号
' key2 : 対象ベクトルのキー列番号
Public Function inner_join(ByVal other As vh_stdvec, _
                     ByVal key1 As Long, _
                     ByVal key2 As Long, _
                     Optional ByRef comp As Variant) As vh_stdvec
    shrink
    Dim other_v As Variant:   other_v = other.clone.free
    If Not IsArray(vec_(0)) Then
        If Not IsArray(other_v(0)) Then
            Set inner_join = inner_join_imple(vec_, _
                                              other_v, _
                                              other_v, _
                                              True, _
                                              comp)
        Else
            Set inner_join = inner_join_imple(vec_, _
                                              mapF(p__n(key2), other_v), _
                                              other_v, _
                                              False, _
                                              comp)
        End If
    Else
        If Not IsArray(other_v(0)) Then
            Set inner_join = inner_join_imple(mapF(p__n(key1), vec_), _
                                              other_v, _
                                              other_v, _
                                              True, _
                                              comp)
        Else
            Set inner_join = inner_join_imple(mapF(p__n(key1), vec_), _
                                              mapF(p__n(key2), other_v), _
                                              other_v, _
                                              False, _
                                              comp)
        End If
    End If
End Function

    Private Function inner_join_imple(ByRef key1_v As Variant, _
                                      ByRef key2_v As Variant, _
                                      ByRef other_v As Variant, _
                                      ByVal same As Boolean, _
                                      ByRef comp As Variant) As vh_stdvec
        Set inner_join_imple = New vh_stdvec
        Dim isCompFunction  As Boolean: isCompFunction = is_bindFun(comp)
        Dim si              As Variant
        If isCompFunction Then
            si = sortIndex_pred(key2_v, comp)
        Else
            si = sortIndex(key2_v)
        End If
        Call permutate(key2_v, si)
        If Not same Then Call permutate(other_v, si)
        Dim i As Long, eq_range As Variant, rec As Variant
        For i = LBound(key1_v) To UBound(key1_v) Step 1
            If isCompFunction Then
                eq_range = equal_range_pred(key2_v, key1_v(i), comp)
            Else
                eq_range = equal_range(key2_v, key1_v(i))
            End If
            If eq_range(0) < eq_range(1) Then
                rec = mapF(p_catV(vec_(i)), subV(other_v, a__o(eq_range(0), eq_range(1))))
                inner_join_imple.push_back_array rec
            End If
        Next i
    End Function
