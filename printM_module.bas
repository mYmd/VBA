Attribute VB_Name = "printM_module"
Option Explicit

'================================================================================
    ' Sub printS                    デバッグウィンドウに配列のサイズを表示する
    ' Sub printM                    デバッグウィンドウに２次元配列を表示する
'================================================================================

'デバッグウィンドウに配列のサイズを表示する
Sub printS(ByRef m As Variant)
Dim mes$, i%, total&
    If VarType(m) = 0 Then Debug.Print " vbEmpty:": Exit Sub
    If Dimension(m) = 0 Then Debug.Print " Scalar": Exit Sub
    mes = "": total = 1
    For i = 1 To Dimension(m) Step 1
        mes = mes & "[Dim" & i & "]: " & LBound(m, i) & " -> " & UBound(m, i) & "  "
        total = total * (1 + UBound(m, i) - LBound(m, i))
    Next i
    mes = mes & ": Total Size = " & total
    Debug.Print mes
End Sub

'デバッグウィンドウに２次元配列を表示する
Sub printM(ByRef m As Variant, Optional r As Variant, Optional c As Variant)
    Dim SR&, ER&, Sc&, EC&, i&, j&, MaxL%(), TMP() As Variant, Msg$
    
    If Dimension(m) = 0 Then Debug.Print m: Exit Sub
    If LBound(m) > UBound(m) Then Debug.Print "#Empty Matrix#": Exit Sub
    If Dimension(m) = 1 Then printV m, r: Exit Sub
    If Dimension(m) > 2 Then Stop: Exit Sub
    If IsMissing(r) = True Then
        SR = LBound(m, 1): ER = UBound(m, 1)
    Else
        If r = 0 Then
            Debug.Print "#Empty Matrix#"
            Exit Sub
        End If
        If r > 0 Then
            SR = LBound(m, 1)
            ER = SR + r - 1
        Else
            SR = UBound(m, 1) + r + 1
            ER = UBound(m, 1)
        End If
    End If
    If IsMissing(c) = True Then
        Sc = LBound(m, 2): EC = UBound(m, 2)
    Else
        If c = 0 Then
            Debug.Print "#Empty Matrix#"
            Exit Sub
        End If
        If c > 0 Then
            Sc = LBound(m, 2)
            EC = Sc + c - 1
        Else
            Sc = UBound(m, 2) + c + 1
            EC = UBound(m, 2)
        End If
    End If
    If SR < LBound(m, 1) Then SR = LBound(m, 1)
    If ER > UBound(m, 1) Then ER = UBound(m, 1)
    If Sc < LBound(m, 2) Then Sc = LBound(m, 2)
    If EC > UBound(m, 2) Then EC = UBound(m, 2)
    If (100000 < (ER - SR + 1) * (EC - Sc + 1)) Then
        Msg = "サイズ超過。縦*横 <=100000以内"
        i = MsgBox(Msg, vbOKOnly, "サイズ超過")
        Exit Sub
    End If
    ReDim MaxL(Sc To EC)
    ReDim TMP(SR To ER, Sc To EC)
    For j = Sc To EC Step 1
        For i = SR To ER Step 1
            TMP(i, j) = m(i, j)
            If IsError(m(i, j)) = True Then TMP(i, j) = "Error!"
            If IsEmpty(m(i, j)) = True Then TMP(i, j) = ""
            If IsNull(m(i, j)) = True Then TMP(i, j) = ""
            If IsArray(m(i, j)) = True Then TMP(i, j) = "[" & i & "," & j & "]"
            If MaxL(j) < LenW(Trim(TMP(i, j))) Then MaxL(j) = LenW(Trim(TMP(i, j)))
        Next i
    Next j
    For i = SR To ER Step 1
        For j = Sc To EC - 1 Step 1
            Debug.Print Spc(2 + MaxL(j) - LenW(Trim(TMP(i, j)))); Trim(TMP(i, j));
        Next j
        Debug.Print Spc(2 + MaxL(UBound(TMP, 2)) - LenW(Trim(TMP(i, UBound(TMP, 2))))); Trim(TMP(i, UBound(TMP, 2)))
    Next i
End Sub
    
'デバッグウィンドウにベクトルを表示する
Private Sub printV(v As Variant, Optional r As Variant)
    Dim SR&, ER&, i&, Msg$
    
    If Dimension(v) = 0 Then Debug.Print v: Exit Sub
    If Dimension(v) = 2 Then printM v, r: Exit Sub
    If LBound(v) > UBound(v) Then Debug.Print "#Empty Vector#": Exit Sub
    If IsMissing(r) = True Then
        SR = LBound(v): ER = UBound(v)
    Else
        If r = 0 Then Debug.Print "#Empty Vector#": Exit Sub
        If r > 0 Then SR = LBound(v): ER = SR + r - 1 Else SR = UBound(v) + r + 1: ER = UBound(v)
    End If
    If SR < LBound(v) Then SR = LBound(v)
    If ER > UBound(v) Then ER = UBound(v)
    If (10000 < ER - SR + 1) Then
        Msg = "サイズ超過。長さ 10000個以内。"
        i = MsgBox(Msg, vbOKOnly, "サイズ超過")
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
            Debug.Print Spc(2); Trim(v(i));
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
        Debug.Print Spc(2); Trim(v(ER))
    End If
End Sub

Private Function LenW(ByRef s As String) As Long
    LenW = LenB(StrConv(s, vbFromUnicode))
End Function