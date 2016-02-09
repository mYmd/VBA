Attribute VB_Name = "Haskell_3_printM"
'Haskell_3_printM
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'================================================================================
'   Sub         printS          デバッグウィンドウに配列のサイズを表示する
'   Sub         printM          デバッグウィンドウに２次元配列を表示する
'   Function    dumpFun         ネストした関数を文字列化
'   Sub         printM_         デバッグウィンドウに1次元ジャグ配列を展開して表示する
'================================================================================

'デバッグウィンドウに配列のサイズを表示する
Sub printS(ByRef m As Variant)
    If IsEmpty(m) Then Debug.Print " vbEmpty:": Exit Sub
    If Dimension(m) = 0 Then Debug.Print " Scalar": Exit Sub
    Dim expr As String, i As Long, total As Long
    expr = "": total = 1
    For i = 1 To Dimension(m) Step 1
        expr = expr & "[Dim" & i & "]: " & LBound(m, i) & " -> " & UBound(m, i) & "  "
        total = total * (1 + UBound(m, i) - LBound(m, i))
    Next i
    expr = expr & ": Total Size = " & total
    Debug.Print expr
End Sub

'デバッグウィンドウに２次元配列を表示する
Sub printM(ByRef m As Variant, Optional ByVal R As Variant, Optional ByVal c As Variant)
    If Dimension(m) = 0 Then Debug.Print m:                     Exit Sub
    If UBound(m) < LBound(m) Then Debug.Print "#Empty Matrix#": Exit Sub
    If Dimension(m) = 1 Then printV m, R:                       Exit Sub
    If 2 < Dimension(m) Then Debug.Print "#Dimension Error#":   Exit Sub
    '-----------------------------------------
    Dim SR As Long, ER As Long
    Dim SC As Long, EC As Long
    Call get_start_end(m, IIf(IsMissing(R), rowSize(m), R), 1, SR, ER)
    Call get_start_end(m, IIf(IsMissing(c), colSize(m), c), 2, SC, EC)
    If (ER < SR) Or (EC < SC) Then
        Debug.Print "#Empty Matrix#"
        Exit Sub
    End If
    Dim i As Long, j As Long
    If (100000 < (ER - SR + 1) * (EC - SC + 1)) Then
        i = MsgBox("サイズ超過。縦*横 <=100000以内", vbOKOnly, "サイズ超過")
        Exit Sub
    End If
    Dim MaxL() As Long:     ReDim MaxL(SC To EC)
    Dim tmp() As Variant:   ReDim tmp(SR To ER, SC To EC)
    For j = SC To EC Step 1
        For i = SR To ER Step 1
            tmp(i, j) = m(i, j)
            If IsError(m(i, j)) = True Then tmp(i, j) = "Error!"
            If IsEmpty(m(i, j)) = True Then tmp(i, j) = ""
            If IsNull(m(i, j)) = True Then tmp(i, j) = ""
            If IsArray(m(i, j)) = True Then tmp(i, j) = "[" & i & "," & j & "]"
            If MaxL(j) < LenW(Trim(tmp(i, j))) Then MaxL(j) = LenW(Trim(tmp(i, j)))
        Next i
    Next j
    For i = SR To ER Step 1
        For j = SC To EC - 1 Step 1
            If VarType(tmp(i, j)) = vbString Then
                Debug.Print Space(2); Trim(tmp(i, j)); Space(MaxL(j) - LenW(Trim(tmp(i, j))));
            Else
                Debug.Print Space(2 + MaxL(j) - LenW(Trim(tmp(i, j)))); Trim(tmp(i, j));
            End If
        Next j
        If VarType(tmp(i, UBound(tmp, 2))) = vbString Then
            Debug.Print Space(2); Trim(tmp(i, UBound(tmp, 2))); Space(MaxL(UBound(tmp, 2)) - LenW(Trim(tmp(i, UBound(tmp, 2)))))
        Else
            Debug.Print Space(2 + MaxL(UBound(tmp, 2)) - LenW(Trim(tmp(i, UBound(tmp, 2))))); Trim(tmp(i, UBound(tmp, 2)))
        End If
    Next i
End Sub
    
    'デバッグウィンドウにベクトルを表示する
    Private Sub printV(ByRef v As Variant, Optional ByVal R As Variant)
        If Dimension(v) = 0 Then Debug.Print v:                     Exit Sub
        If Dimension(v) = 2 Then printM v, R:                       Exit Sub
        If LBound(v) > UBound(v) Then Debug.Print "#Empty Vector#": Exit Sub
        '-----------------------------------------
        Dim SR As Long, ER As Long
        Call get_start_end(v, IIf(IsMissing(R), sizeof(v), R), 1, SR, ER)
        If ER < SR Then
            Debug.Print "#Empty Vector#"
            Exit Sub
        End If
        Dim i As Long
        If (10000 < ER - SR + 1) Then
            i = MsgBox("サイズ超過。長さ 10000個以内。", vbOKOnly, "サイズ超過")
            Exit Sub
        End If
        For i = SR To ER - 1 Step 1
            If IsError(v(i)) = True Then
                Debug.Print "  Error!";
            ElseIf IsArray(v(i)) = True Then
                Debug.Print "  [" & i & "]";
            ElseIf IsEmpty(v(i)) = True Then
                Debug.Print "  ";
            ElseIf IsNull(v(i)) = True Then
                Debug.Print "  ";
            Else
                Debug.Print Space(2); Trim(v(i));
            End If
        Next i
        If IsError(v(ER)) = True Then
            Debug.Print "  Error!"
        ElseIf IsArray(v(ER)) = True Then
            Debug.Print "  [" & ER & "]"
        ElseIf IsEmpty(v(ER)) = True Then
            Debug.Print "  "
        ElseIf IsNull(v(ER)) = True Then
            Debug.Print "  "
        Else
            Debug.Print Space(2); Trim(v(ER))
        End If
    End Sub

    Private Function LenW(ByRef s As String) As Long
        LenW = LenB(StrConv(s, vbFromUnicode))
    End Function

    Private Sub get_start_end(ByRef m As Variant, _
                              ByVal length As Long, _
                              ByVal dimen As Long, _
                              ByRef start_ As Long, _
                              ByRef end_ As Long)
        If length > 0 Then
            start_ = LBound(m, dimen)
            end_ = start_ + length - 1
            If UBound(m, dimen) < end_ Then end_ = UBound(m, dimen)
        Else
            start_ = UBound(m, dimen) + length + 1
            end_ = UBound(m, dimen)
            If start_ < LBound(m, dimen) Then start_ = LBound(m, dimen)
        End If
    End Sub

'ネストした関数を文字列化
Function dumpFun(ByRef x As Variant, Optional ByVal OneTwo As Long = 0) As Variant
    If is_bindFun(x) Then
        dumpFun = "F" & (x(0) Mod 10000) & _
                  "(" & dumpFun(x(1), IIf(OneTwo = 0, 1, OneTwo)) & _
                  ", " & dumpFun(x(2), IIf(OneTwo = 0, 2, OneTwo)) & ")"
    ElseIf is_placeholder(x) Then
        If x = placeholder(0) Then
            If OneTwo = 1 Then
                dumpFun = "_1"
            ElseIf OneTwo = 2 Then
                dumpFun = "_2"
            Else
                dumpFun = "_"
            End If
        ElseIf x = placeholder(1) Then
            dumpFun = "_1"
        ElseIf x = placeholder(2) Then
            dumpFun = "_2"
        Else
            dumpFun = ""
        End If
    ElseIf IsEmpty(x) Then
        dumpFun = ""
    Else
        If VarType(x) = vbString Then
            dumpFun = """" & x & """"
        ElseIf IsNumeric(x) Then
            dumpFun = x
        Else
            dumpFun = "*"
        End If
    End If
End Function

'デバッグウィンドウに1次元ジャグ配列を展開して表示する
Sub printM_(ByRef vec As Variant, Optional ByVal R As Variant)
    Dim begin_ As Long, end_   As Long
    If IsMissing(R) Then
        begin_ = LBound(vec)
        end_ = UBound(vec) + 1
    ElseIf 0 < R Then
        begin_ = LBound(vec)
        end_ = min_fun(UBound(vec) + 1, begin_ + R)
    Else
        end_ = UBound(vec) + 1
        begin_ = max_fun(LBound(vec), end_ + R)
    End If
    Do While begin_ < end_
        printM vec(begin_)
        begin_ = begin_ + 1
    Loop
End Sub
