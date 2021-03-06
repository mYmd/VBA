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
Sub printS(ParamArray m() As Variant)
    Dim i As Long
    For i = LBound(m) To UBound(m) Step 1
        printS_imple m(i)
    Next i
End Sub
    Private Sub printS_imple(ByRef m As Variant)
        If IsEmpty(m) Then Debug.Print " vbEmpty:": Exit Sub
        If Dimension(m) = 0 Then
            If IsArray(m) Then
                Debug.Print "#Erased Array#":   Exit Sub
            Else
                Debug.Print " Scalar":          Exit Sub
            End If
        End If
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
Sub printM(ByRef m As Variant, Optional ByRef r As Variant, Optional ByRef C As Variant)
    If Dimension(m) = 0 Then
        If IsArray(m) Then
            Debug.Print "#Erased Array#":                       Exit Sub
        ElseIf IsObject(m) Then
            Debug.Print " " & TypeName(m):                      Exit Sub
        Else
            Debug.Print m:                                      Exit Sub
        End If
    End If
    If UBound(m) < LBound(m) Then Debug.Print "#Empty Matrix#": Exit Sub
    If Dimension(m) = 1 Then printV m, r:                       Exit Sub
    If 2 < Dimension(m) Then Debug.Print "#Dimension Error#":   Exit Sub
    '-----------------------------------------
    Dim SR As Long, er As Long
    Dim SC As Long, EC As Long
    Call get_start_end(m, IIf(IsMissing(r), rowSize(m), r), 1, SR, er)
    Call get_start_end(m, IIf(IsMissing(C), colSize(m), C), 2, SC, EC)
    If (er < SR) Or (EC < SC) Then
        Debug.Print "#Empty Matrix#"
        Exit Sub
    End If
    Dim i As Long, j As Long
    If (100000 < (er - SR + 1) * (EC - SC + 1)) Then
        Debug.Print "# The matrix is too big to display!!!! (>100000) #"
        Exit Sub
    End If
    Dim MaxL() As Long:     ReDim MaxL(SC To EC)
    Dim tmp() As Variant:   ReDim tmp(SR To er, SC To EC)
    For j = SC To EC Step 1
        For i = SR To er Step 1
            If IsObject(m(i, j)) = True Then
                tmp(i, j) = TypeName(m(i, j))
            Else
                tmp(i, j) = m(i, j)
                If IsError(m(i, j)) = True Then tmp(i, j) = "Error!"
                If IsEmpty(m(i, j)) = True Then tmp(i, j) = ""
                If IsNull(m(i, j)) = True Then tmp(i, j) = ""
                If IsArray(m(i, j)) = True Then tmp(i, j) = "[" & i & "," & j & "]"
                If IsObject(m(i, j)) = True Then tmp(i, j) = TypeName(m(i, j))
            End If
            If MaxL(j) < LenW(Trim(tmp(i, j))) Then MaxL(j) = LenW(Trim(tmp(i, j)))
        Next i
    Next j
    For i = SR To er Step 1
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
    Private Sub printV(ByRef v As Variant, Optional ByRef r As Variant)
        If Dimension(v) = 0 Then Debug.Print v:                     Exit Sub
        If Dimension(v) = 2 Then printM v, r:                       Exit Sub
        If LBound(v) > UBound(v) Then Debug.Print "#Empty Vector#": Exit Sub
        '-----------------------------------------
        Dim SR As Long, er As Long
        If IsMissing(r) Then
            Call get_start_end(v, sizeof(v), 1, SR, er)
        ElseIf rowSize(r) < 2 Then
            Call get_start_end(v, r, 1, SR, er)
        Else
            SR = r(0): er = r(1)
        End If
        If er < SR Then
            Debug.Print "#Empty Vector#"
            Exit Sub
        End If
        Dim i As Long
        If (10000 < er - SR + 1) Then
            Debug.Print "# The vector is too big to display!!!! (>10000) #"
            Exit Sub
        End If
        For i = SR To er - 1 Step 1
            If IsError(v(i)) = True Then
                Debug.Print "  Error!";
            ElseIf IsArray(v(i)) = True Then
                Debug.Print "  [" & i & "]";
            ElseIf IsEmpty(v(i)) = True Then
                Debug.Print "  ";
            ElseIf IsNull(v(i)) = True Then
                Debug.Print "  ";
            ElseIf IsObject(v(i)) = True Then
                Debug.Print Space(2); TypeName(v(i));
            Else
                Debug.Print Space(2); Trim(v(i));
            End If
        Next i
        If IsError(v(er)) = True Then
            Debug.Print "  Error!"
        ElseIf IsArray(v(er)) = True Then
            Debug.Print "  [" & er & "]"
        ElseIf IsEmpty(v(er)) = True Then
            Debug.Print "  "
        ElseIf IsNull(v(er)) = True Then
            Debug.Print "  "
        ElseIf IsObject(v(er)) = True Then
            Debug.Print Space(2); TypeName(v(er))
        Else
            Debug.Print Space(2); Trim(v(er))
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
        If x = ph_0 Then
            If OneTwo = 1 Then
                dumpFun = "!1"
            ElseIf OneTwo = 2 Then
                dumpFun = "!2"
            Else
                dumpFun = "!"
            End If
        ElseIf x = ph_1 Then
            dumpFun = "!1"
        ElseIf x = ph_2 Then
            dumpFun = "!2"
        ElseIf x = yield_0 Then
            If OneTwo = 1 Then
                dumpFun = "_1"
            ElseIf OneTwo = 2 Then
                dumpFun = "_2"
            Else
                dumpFun = "_"
            End If
        ElseIf x = yield_1 Then
            dumpFun = "_1"
        ElseIf x = yield_2 Then
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
Sub printM_(ByRef vec As Variant, Optional ByRef r As Variant, Optional ByRef C As Variant)
    Select Case Dimension(vec)
    Case 0
        printM vec
    Case 1
        Dim begin_ As Long, end_   As Long
        If IsMissing(r) Then
            begin_ = LBound(vec)
            end_ = UBound(vec)
        ElseIf 1 < rowSize(r) Then
            begin_ = r(0)
            end_ = r(1)
        ElseIf 0 <= r Then
            begin_ = LBound(vec)
            end_ = min_fun(UBound(vec), begin_ + r - 1)
        Else
            end_ = UBound(vec)
            begin_ = max_fun(LBound(vec), end_ + r + 1)
        End If
        Do While begin_ <= end_
            printM vec(begin_), C
            begin_ = begin_ + 1
        Loop
    End Select
End Sub
