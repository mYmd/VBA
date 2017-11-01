Attribute VB_Name = "misc_utility"
'misc_utility
'Copyright (c) 2016 mmYYmmdd
Option Explicit

'*********************************************************************************
'   ���[�e�B���e�B
'*********************************************************************************
'   Function  p__n                      p_getNth_b(, n)�̍\����
'   Function  p_try                     IIf(pred(a), a', b')
'   Function  p_try_not                 IIf(Not pred(a), a', b')�̍\����
'   Function  p_try_less                p_try(p_less(p__n(0), p__n(1)), p__n(0), Null) �̍\����
'   Function  p_vartype                 VarType�l
'   Function  p_typename                �f�[�^�^��
'   Function  p_isNumeric               IsNumeric�֐�
'   Function  p_format                  Format�֐�
'   Function  p_InStr                   InStr�֐�
'   Function  p_InStrRev                InStrRev�֐�
'   Function  p_Like                    Like�֐�
'   Function  p_StrConv                 StrConv�֐�
'   Function  p_Trim                    Trim�֐�
'   Function  separate_string           ������̍��E����
'   Function  subM_R                    subM(m, �s�͈�) �̍\����
'   Function  subM_R_b                  �V�i�I�t�Z�b�g�A�h���X�j
'   Function  subM_C                    subM(m, , ��͈�) �̍\����
'   Function  subM_C_b                  �V�i�I�t�Z�b�g�A�h���X�j
'   Function  selectRow_b               selectRow�i�I�t�Z�b�g�A�h���X�j
'   Function  selectCol_b               selectCol�i�I�t�Z�b�g�A�h���X�j
'   Sub       fillRow_b                 fillRow�i�I�t�Z�b�g�A�h���X�j
'   Function  fillRow_b_move            fillRow_move�i�I�t�Z�b�g�A�h���X�j
'   Sub       fillCol_b                 fillCol�i�I�t�Z�b�g�A�h���X�j
'   Function  fillCol_b_move            ��fillCol_move�i�I�t�Z�b�g�A�h���X�j
'   Function  adjacent_op               1�����z��̗אڂ���v�f�Ԃ�2������
'   Function  get_unique                1�����z��̏d���v�f���폜���� (�\�[�g�ϑO��)
'  -----------------------------------------------------------------------------
'   Sub       rowWise_change            2�����z��̍s���ƂɊ֐��K�p
'   Function  rowWise_change_move       �Vmove���ĕԂ�
'   Sub       columnWise_change         2�����z��̗񂲂ƂɊ֐��K�p
'   Function  columnWise_change_move    �Vmove���ĕԂ�
'   Function  equal_all                 1�����z��̑S�v�f�̓��l��r
'   Function  equal_all_pred            �V�@�q��o�[�W����
'   Function  filter_if                 �q���^����1�����z����t�B���^�����O
'   Function  filter_if_not             �q���^����1�����z����t�B���^�����O�i�ے�`�j
'  -----------------------------------------------------------------------------
'   Function  p_not                     �_��Not
'   Function  p_imply                   �܈�(A=>B)   IIF(Not A Or B, True, False)
'  -----------------------------------------------------------------------------
'   Function  pipe                      vh_pipe�I�u�W�F�N�g�̐���
'   Function  pipe_                     vh_pipe�I�u�W�F�N�g�̐����i������move����j
'  -----------------------------------------------------------------------------
'   Function  stdVec                    vh_stdvec�I�u�W�F�N�g�̐���
'  -----------------------------------------------------------------------------
'   Function  splitStr2Funs             delimiter�ŋ�؂�ꂽ��������֐���փ}�b�s���O
'   Function  str2SummaryFun            �����񂩂�W�v�֐��֕ϊ�
'   Function  str2ConvertFun            �����񂩂�^�ϊ��֐��֕ϊ�
'  -----------------------------------------------------------------------------
'   Function  group_by_partition_points partition_points �ɂ��GROUP-BY
'  -----------------------------------------------------------------------------
'   Function  group_by_pred             �������Ƃ̔z��ɕ���
'  -----------------------------------------------------------------------------
'   Function  csv2Vector                csv�t�@�C����1�s��z��ɕ���
'  -----------------------------------------------------------------------------
'   Function  A_overlap_B               �Q��1�����z��̋��L�����Ɣ񋤗L���������
'*********************************************************************************

' p_getNth_b(, n)�̍\����
Public Function p__n(ByVal n As Long) As Variant
    p__n = p_getNth_b(, n)
End Function

' IIf(pred(a), a', b')�̍\����
Public Function p_try(ByRef pred As Variant, _
                        Optional ByRef f1 As Variant, Optional ByRef f2 As Variant) As Variant
    If IsMissing(f1) Then
        If IsMissing(f2) Then
            p_try = p_replace_0(p_if_else(, VBA.Array(pred, p_makeSole, 0)))
        Else
            p_try = p_replace_0(p_if_else(, VBA.Array(pred, p_makeSole, 0)), f2)
        End If
    Else
        If IsMissing(f2) Then
            p_try = p_replace_0(p_if_else(, VBA.Array(pred, p_makeSole(f1), 0)))
        Else
            p_try = p_replace_0(p_if_else(, VBA.Array(pred, p_makeSole(f1), 0)), f2)
        End If
    End If
End Function

' IIf(Not pred(a), a', b')�̍\����
Public Function p_try_not(ByRef pred As Variant, _
                        Optional ByRef f1 As Variant, Optional ByRef f2 As Variant) As Variant
    If IsMissing(f1) Then
        If IsMissing(f2) Then
            p_try_not = p_replace_0(p_if_else(, VBA.Array(pred, 0, p_makeSole)))
        Else
            p_try_not = p_replace_0(p_if_else(, VBA.Array(pred, 0, p_makeSole)), f2)
        End If
    Else
        If IsMissing(f2) Then
            p_try_not = p_replace_0(p_if_else(, VBA.Array(pred, 0, p_makeSole(f1))))
        Else
            p_try_not = p_replace_0(p_if_else(, VBA.Array(pred, 0, p_makeSole(f1))), f2)
        End If
    End If
End Function
    
    Private Function replace_0(ByRef x As Variant, ByRef alt As Variant) As Variant
        If IsNumeric(x) Then
            Call assignVar(replace_0, alt)
        Else
            Call assignVar(replace_0, x(0))
        End If
    End Function
    Private Function p_replace_0(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_replace_0 = make_funPointer(AddressOf replace_0, firstParam, secondParam)
    End Function

' p_try(p_less(p__n(0), p__n(1)), p__n(0), Null) �̍\����
' equal_range �̒l�� subV_if �ɑ������Ƃ����ɕ֗�
Public Function p_try_less()
    p_try_less = p_try(p_less(p__n(0), p__n(1)), p__n(0), Null)
End Function

' VarType�l
    Private Function vartype_(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
        vartype_ = VarType(x)
    End Function
Public Function p_vartype(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_vartype = make_funPointer(AddressOf vartype_, firstParam, secondParam)
End Function

' �f�[�^�^��
    Private Function typename_(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
        typename_ = TypeName(x)
    End Function
Public Function p_typename(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_typename = make_funPointer(AddressOf typename_, firstParam, secondParam)
End Function

' IsNumeric�֐�
        Private Function IsNumeric_(ByRef expr As Variant, Optional ByRef strict As Variant) As Variant
            IsNumeric_ = IIf(IsNumeric(expr) Or IsDate(expr), 1&, 0&)
            If Not IsMissing(strict) Then
                If 0 = Not_(strict) Then
                    IsNumeric_ = IIf(IsNumeric(expr) And Not IsEmpty(expr) Or VarType(expr) <> vbString, 0&, 1&)
                End If
            End If
        End Function
Public Function p_isNumeric(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_isNumeric = make_funPointer_with_2nd_Default(AddressOf IsNumeric_, firstParam, secondParam)
End Function

' Format�֐�
    Private Function format_(ByRef expr As Variant, ByRef fmt As Variant) As Variant
        format_ = Format(expr, fmt)
    End Function
Public Function p_format(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_format = make_funPointer(AddressOf format_, firstParam, secondParam)
End Function

' InStr�֐�
     Private Function InStr_(ByRef s As Variant, ByRef expr As Variant) As Variant
        InStr_ = InStr(s, expr)
        If IsNull(InStr_) Then InStr_ = 0
    End Function
Public Function p_InStr(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_InStr = make_funPointer(AddressOf InStr_, firstParam, secondParam)
End Function

' InStrRev�֐�
    Private Function InStrRev_(ByRef s As Variant, ByRef expr As Variant) As Variant
        InStrRev_ = InStrRev(s, expr)
        If IsNull(InStrRev_) Then InStrRev_ = 0
    End Function
Public Function p_InStrRev(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_InStrRev = make_funPointer(AddressOf InStrRev_, firstParam, secondParam)
End Function

' Like�֐�
    Private Function Like_(ByRef s As Variant, ByRef expr As Variant) As Variant
        Like_ = IIf(s Like expr, 1, 0)
    End Function
Public Function p_Like(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_Like = make_funPointer(AddressOf Like_, firstParam, secondParam)
End Function

' StrConv�֐�
     Private Function StrConv_(ByRef s As Variant, ByRef expr As Variant) As Variant
        StrConv_ = StrConv(s, expr)
     End Function
Public Function p_StrConv(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_StrConv = make_funPointer(AddressOf StrConv_, firstParam, secondParam)
End Function

' Trim�֐�
     Private Function Trim_(ByRef expr As Variant, Optional ByRef left_right As Variant) As Variant
        If IsNumeric(left_right) Then
            If left_right < 0 Then
                Trim_ = RTrim(expr)
            ElseIf 0 < left_right Then
                Trim_ = LTrim(expr)
            Else
                Trim_ = Trim(expr)
            End If
        Else
            Trim_ = Trim(expr)
        End If
     End Function
Public Function p_Trim(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_Trim = make_funPointer_with_2nd_Default(AddressOf Trim_, firstParam, secondParam)
End Function

' ������̍��E����
Function separate_string(ByRef expr As Variant, ByRef n As Variant) As Variant
    If 0 < n Then
        separate_string = VBA.Array(Left(expr, n), str_right(expr, -n))
    Else
        separate_string = VBA.Array(str_left(expr, n), Right(expr, -n))
    End If
End Function
    Function p_separate_string(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_separate_string = make_funPointer(AddressOf separate_string, firstParam, secondParam)
    End Function

' subM(m, �s�͈�) �̍\����
Public Function subM_R(ByRef m As Variant, ByRef rRange As Variant) As Variant
    If IsArray(rRange) Then
        subM_R = subM(m, rRange)
    Else
        subM_R = subM(m, VBA.Array(rRange))
    End If
End Function
    Public Function p_subM_R(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subM_R = make_funPointer(AddressOf subM_R, firstParam, secondParam)
    End Function

' subM(m, �s�͈�) �̍\�����i�I�t�Z�b�g�A�h���X�j
Public Function subM_R_b(ByRef m As Variant, ByRef rRange As Variant) As Variant
    Dim range_b As Variant
    range_b = mapF(p_if_else(, Array(p_less_equal(0), p_plus(LBound(m, 1)), p_plus(1 + UBound(m, 1)))), rRange)
    subM_R_b = subM_R(m, range_b)
End Function
    Public Function p_subM_R_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subM_R_b = make_funPointer(AddressOf subM_R_b, firstParam, secondParam)
    End Function

' subM(m, , ��͈�) �̍\����
Public Function subM_C(ByRef m As Variant, ByRef cRange As Variant) As Variant
    If IsArray(cRange) Then
        subM_C = subM(m, , cRange)
    Else
        subM_C = subM(m, , VBA.Array(cRange))
    End If
End Function
    Public Function p_subM_C(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subM_C = make_funPointer(AddressOf subM_C, firstParam, secondParam)
    End Function

' subM(m, , ��͈�) �̍\�����i�I�t�Z�b�g�A�h���X�j
Public Function subM_C_b(ByRef m As Variant, ByRef cRange As Variant) As Variant
    Dim range_b As Variant
    range_b = mapF(p_if_else(, Array(p_less_equal(0), p_plus(LBound(m, 2)), p_plus(1 + UBound(m, 2)))), cRange)
    subM_C_b = subM_C(m, range_b)
End Function
    Public Function p_subM_C_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_subM_C_b = make_funPointer(AddressOf subM_C_b, firstParam, secondParam)
    End Function

'����s�̎擾�i�I�t�Z�b�g�A�h���X�j
'index < 0 �̏ꍇ�͌�납��擾
Public Function selectRow_b(ByRef matrix As Variant, ByRef i As Variant) As Variant
    If 0 <= i Then
        selectRow_b = selectRow(matrix, i + LBound(matrix, 1))
    Else
        selectRow_b = selectRow(matrix, i + 1 + UBound(matrix, 1))
    End If
End Function
    Public Function p_selectRow_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_selectRow_b = make_funPointer(AddressOf selectRow_b, firstParam, secondParam)
    End Function

'�����̎擾�i�I�t�Z�b�g�A�h���X�j
'index < 0 �̏ꍇ�͌�납��擾
Public Function selectCol_b(ByRef matrix As Variant, ByRef j As Variant) As Variant
    If 0 <= j Then
        selectCol_b = selectCol(matrix, j + LBound(matrix, 2))
    Else
        selectCol_b = selectCol(matrix, j + 1 + UBound(matrix, 2))
    End If
End Function
    Public Function p_selectCol_b(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_selectCol_b = make_funPointer(AddressOf selectCol_b, firstParam, secondParam)
    End Function

'�z��̓���s���f�[�^�Ŗ��߂�i�I�t�Z�b�g�A�h���X�j
Public Sub fillRow_b(ByRef matrix As Variant, ByVal i As Long, ByRef data As Variant)
    If 0 <= i Then
        Call fillRow(matrix, i + LBound(matrix, 1), data)
    Else
        Call fillRow(matrix, i + 1 + UBound(matrix, 1), data)
    End If
End Sub

'�z��̓���s���f�[�^�Ŗ��߂�move���ĕԂ��i�I�t�Z�b�g�A�h���X�j
Public Function fillRow_b_move(ByRef matrix As Variant, ByVal i As Long, ByRef data As Variant) As Variant
    Call fillRow_b(matrix, i, data)
    fillRow_b_move = moveVariant(matrix)
End Function

'�z��̓������f�[�^�Ŗ��߂�i�I�t�Z�b�g�A�h���X�j
Public Sub fillCol_b(ByRef matrix As Variant, ByVal j As Long, ByRef data As Variant)
    If 0 <= j Then
        Call fillCol(matrix, j + LBound(matrix, 2), data)
    Else
        Call fillCol(matrix, j + 1 + UBound(matrix, 2), data)
    End If
End Sub

'�z��̓������f�[�^�Ŗ��߂�move���ĕԂ��i�I�t�Z�b�g�A�h���X�j
Public Function fillCol_b_move(ByRef matrix As Variant, ByVal j As Long, ByRef data As Variant) As Variant
    Call fillCol_b(matrix, j, data)
    fillCol_b_move = moveVariant(matrix)
End Function

'*********************************************************************************
' 1�����z��vec�̗אڂ���v�f�Ԃ�2������op���s��
' �o�͗�̗v�f���͌��̗v�f�� - 1   (LBound = 0)
Public Function adjacent_op(ByRef op As Variant, ByRef vec As Variant) As Variant
    adjacent_op = VBA.Array()
    If is_bindFun(op) And Dimension(vec) = 1 And 1 < sizeof(vec) Then
        Dim ret As Variant
        ret = self_zipWith(op, vec, 1)
        ReDim Preserve ret(0 To UBound(ret) - 1)
        swapVariant adjacent_op, ret
    End If
End Function
    Public Function p_adjacent_op(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_adjacent_op = make_funPointer(AddressOf adjacent_op, firstParam, secondParam, 1)
    End Function

' �z��̏d���v�f���폜����(�\�[�g�ϑO��Acomp�͓��l����)
Public Function get_unique(ByRef vec As Variant, Optional ByRef comp As Variant) As Variant
    Dim flag As Variant
    If IsMissing(comp) Then
        flag = self_zipWith(p_notEqual, vec, -1)
    Else
        flag = mapF(p_equal(0), self_zipWith(comp, vec, -1))
    End If
    If 0 < sizeof(flag) Then flag(0) = 1
    get_unique = filterR(vec, flag)
End Function
    Function p_get_unique(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_get_unique = make_funPointer_with_2nd_Default(AddressOf get_unique, firstParam, secondParam, 2)
    End Function

' 2�����z��̍s���ƂɊ֐��K�p
Public Sub rowWise_change(ByRef matrix As Variant, ByRef funcs As Variant)
    Dim i As Long
    For i = 0 To min_fun(rowSize(matrix), sizeof(funcs)) - 1 Step 1
        If is_bindFun(getNth_b(funcs, i)) Then
            Call fillRow_b(matrix, i, mapF(getNth_b(funcs, i), selectRow_b(matrix, i)))
        End If
    Next i
End Sub

' 2�����z��̍s���ƂɊ֐��K�p��move���ĕԂ�
Public Function rowWise_change_move(ByRef matrix As Variant, ByRef funcs As Variant) As Variant
    Call rowWise_change(matrix, funcs)
    rowWise_change_move = moveVariant(matrix)
End Function
    Public Function p_rowWise_change_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_rowWise_change_move = make_funPointer(AddressOf rowWise_change_move, firstParam, secondParam)
    End Function

' 2�����z��̗񂲂ƂɊ֐��K�p
Public Sub columnWise_change(ByRef matrix As Variant, ByRef funcs As Variant)
    Dim j As Long
    For j = 0 To min_fun(colSize(matrix), sizeof(funcs)) - 1 Step 1
        If is_bindFun(getNth_b(funcs, j)) Then
            Call fillCol_b(matrix, j, mapF(getNth_b(funcs, j), selectCol_b(matrix, j)))
        End If
    Next j
End Sub

' 2�����z��̗񂲂ƂɊ֐��K�p��move���ĕԂ�
Public Function columnWise_change_move(ByRef matrix As Variant, ByRef funcs As Variant) As Variant
    Call columnWise_change(matrix, funcs)
    columnWise_change_move = moveVariant(matrix)
End Function
    Public Function p_columnWise_change_move(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_columnWise_change_move = make_funPointer(AddressOf columnWise_change_move, firstParam, secondParam)
    End Function

' 1�����z��̑S�v�f�̓��l��r
Public Function equal_all(ByRef a As Variant, ByRef b As Variant) As Variant
    equal_all = equal_all_pred(p_equal, a, b)
End Function
    Public Function p_equal_all(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_equal_all = make_funPointer(AddressOf equal_all, firstParam, secondParam)
    End Function

' 1�����z��̑S�v�f�̓��l��r�i�q��o�[�W�����j
Public Function equal_all_pred(ByRef pred As Variant, ByRef a As Variant, ByRef b As Variant) As Variant
    If sizeof(a) = sizeof(b) Then
        equal_all_pred = IIf(sizeof(a) <= find_pred(p_equal(0), zipWith(pred, a, b)), _
                             1, _
                             0)
    Else
        equal_all_pred = 0
    End If
End Function

' �q���^����1�����z����t�B���^�����O
Public Function filter_if(ByRef fun As Variant, ByRef vec As Variant) As Variant
    filter_if = filterR(vec, mapF(fun, vec))
End Function
    Public Function p_filter_if(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_filter_if = make_funPointer(AddressOf filter_if, firstParam, secondParam, 1)
    End Function

' �q���^����1�����z����t�B���^�����O�i�ے�`�j
Public Function filter_if_not(ByRef fun As Variant, ByRef vec As Variant) As Variant
    filter_if_not = filter_if(p_equal(0, fun), vec)
End Function
    Public Function p_filter_if_not(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_filter_if_not = make_funPointer(AddressOf filter_if_not, firstParam, secondParam, 1)
    End Function

' �_��Not   (0, Null, Empty, Nothing, CDate(0) ��False�Ƃ݂Ȃ�)
     Private Function Not_(ByRef expr As Variant, Optional ByRef dummy As Variant) As Variant
        If IsNull(expr) Then
            Not_ = 1
        ElseIf IsEmpty(expr) Then
            Not_ = 1
        ElseIf IsNumeric(expr) Then
            Not_ = IIf(expr = False, 1, 0)
        ElseIf IsObject(expr) Then
            Not_ = IIf(expr Is Nothing, 1, 0)
        ElseIf IsDate(expr) Then
            Not_ = IIf(expr = 0, 1, 0)
        Else
            Not_ = 0
        End If
    End Function
Public Function p_not(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_not = make_funPointer(AddressOf Not_, firstParam, secondParam)
End Function

' �܈�(A=>B)   IIF(Not A Or B, True, False)
    Private Function Imply_(ByRef a As Variant, ByRef b As Variant) As Variant
        Imply_ = IIf(Not_(a) = 1 Or Not_(b) = 0, 1, 0)
    End Function
Public Function p_imply(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_imply = make_funPointer(AddressOf Imply_, firstParam, secondParam)
End Function

'-----------------------------------------------------------
' vh_pipe�I�u�W�F�N�g�̐���
Public Function pipe(ByRef x As Variant) As vh_pipe
    Set pipe = New vh_pipe
    pipe.swap_val_ (x)   '�R�s�[��n��
End Function

' vh_pipe�I�u�W�F�N�g�̐����i������move����j
Public Function pipe_(ByRef x As Variant) As vh_pipe
    Set pipe_ = New vh_pipe
    pipe_.swap_val_ x
End Function

'-----------------------------------------------------------
' vh_stdvec�I�u�W�F�N�g�̐���
Public Function stdVec(Optional ByRef x As Variant) As vh_stdvec
    Set stdVec = New vh_stdvec
    stdVec.from x
End Function


'******************************************************************************
' delimiter�ŋ�؂�ꂽ��������֐���փ}�b�s���O
' strFuns   : �֐���\��������
' my_str2Fun: �����񂩂�֐��ւ̃}�b�s���O�֐�
' delimiter : strFuns�̋�؂蕶��
' ��j%f%d%s%n �� Array(f, d, s, n)
'******************************************************************************
Public Function splitStr2Funs(ByVal strFuns As String, _
                              ByRef my_str2Fun As Variant, _
                              ByVal delimiter As String) As Variant
    If Left(strFuns, Len(delimiter)) = delimiter Then
        strFuns = Right(strFuns, Len(strFuns) - Len(delimiter))
    End If
    splitStr2Funs = mapF(my_str2Fun, Split(strFuns, delimiter))
End Function

' �isplitStr2Funs �̑Ώۊ֐��j
' �����񂩂�W�v�֐��֕ϊ�
' �Ǝ��̕ϊ��֐��������Ƃ��͂���Case Else �̒��ł��̊֐����Ăяo���`�ɂ���Ƃ�������
' %t %tp  %top      : �擪
' %b %btm %bottom   : ����
' %c %cnt %count    : ��
' %s %sum %�v       : ���v
' %a %avg %average  : ����
' %max              : �ő�
' %min              : �ŏ�
Public Function str2SummaryFun(ByRef s As Variant, Optional ByRef other As Variant) As Variant
    Select Case StrConv(s, vbNarrow + vbLowerCase)
        Case "t", "tp", "top"
            str2SummaryFun = p_getNth_b(, 0)
        Case "b", "btm", "bottom"
            str2SummaryFun = p_getNth_b(, -1)
        Case "c", "cnt", "count"
            str2SummaryFun = p_sizeof()
        Case "s", "sum", "�v"
            str2SummaryFun = p_foldl1(p_plus(yield_1, yield_2))
        Case "a", "avg", "average"
            str2SummaryFun = p_divide(p_foldl1(p_plus(yield_1, yield_2)), p_sizeof)
        Case "max"
            str2SummaryFun = p_foldl1(p_max(yield_1, yield_2))
        Case "min"
            str2SummaryFun = p_foldl1(p_min(yield_1, yield_2))
        Case ""
            str2SummaryFun = p_constant(other)
        Case Else
            str2SummaryFun = p_constant(s)
    End Select
End Function
    Public Function p_str2SummaryFun(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_str2SummaryFun = make_funPointer(AddressOf str2SummaryFun, firstParam, secondParam)
    End Function

' �isplitStr2Funs �̑Ώۊ֐��j
' �����񂩂�^�ϊ��֐��֕ϊ�
' �Ǝ��̕ϊ��֐��������Ƃ��͂���Case Else �̒��ł��̊֐����Ăяo���`�ɂ���Ƃ�������
' %s*  : Format( ,*)
' %d   : ������
' %f   : ������
' %s   : ������
Public Function str2ConvertFun(ByRef s As Variant, ByRef dummy As Variant) As Variant
    Dim expr As String: expr = StrConv(s, vbNarrow + vbLowerCase)
    If Left(expr, 1) = "s" Then
        If expr = "s" Then
            str2ConvertFun = p_CStr
        Else
            str2ConvertFun = p_format(, Right(s, Len(s) - 1))
        End If
    Else
        Select Case expr
        Case "d"
            str2ConvertFun = p_CLng
        Case "f"
            str2ConvertFun = p_CDbl
        Case Else
            str2ConvertFun = Empty
        End Select
    End If
End Function
    Public Function p_str2ConvertFun(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_str2ConvertFun = make_funPointer(AddressOf str2ConvertFun, firstParam, secondParam)
    End Function

'******************************************************************************
' partition_points �ɂ��GROUP-BY
' matrix    : �Ώ۔z��i2�����z��܂��̓W���O�z��j
' pp        : partition_points �i�W�v����s�͈͂���؂�s�ԍ��̏W���j
' strFuns   : �񂲂Ƃ̏W�v�֐���\��������
' my_str2Fun: �����񂩂�W�v�֐��ւ̃}�b�s���O�֐��istr2SummaryFun���f�t�H���g�j
' ��jgroup_by_partition_points(matrix, pp, "%t%c%s%a%min%max")
'******************************************************************************
Public Function group_by_partition_points(ByRef matrix As Variant, _
                                          ByRef pp As Variant, _
                                          ByRef strFuns As String, _
                                 Optional ByVal my_str2Fun As Variant) As Variant
    If IsMissing(my_str2Fun) Then my_str2Fun = p_str2SummaryFun(, "-")    '�f�t�H���g��
    Dim funs As Variant
    funs = splitStr2Funs(strFuns, my_str2Fun, "%")
    Dim intervals As Variant
    intervals = adjacent_op(p_a__o, pp)
    Dim ranges As Variant
    ranges = mapF_swap(p_subM_R, matrix, intervals)
    group_by_partition_points = unzip(mapF_swap(p_summaryUnit, , funs, ranges), 2)
End Function
    ' �X�̏W�v�s�͈͂̏���
    Private Function summaryUnit(ByRef matrix As Variant, ByRef funs As Variant) As Variant
        Select Case Dimension(matrix)
            Case 1: summaryUnit = zipWith(p_applyFun, unzip(matrix, 1), funs)
            Case 2: summaryUnit = zipWith(p_applyFun, zipR(matrix), funs)
        End Select
    End Function
    Private Function p_summaryUnit(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_summaryUnit = make_funPointer(AddressOf summaryUnit, firstParam, secondParam)
    End Function

'******************************************************************************
'�������Ƃ̔z��ɕ�������iv��move�����j
Function group_by_pred(ByRef v As Variant, ByRef pred As Variant) As Variant
    Dim prop As Variant, s As Variant, pp As Variant, ret As Variant
    Call changeLBound(v, 0)
    prop = mapF(pred, v)
    s = sortIndex(prop)
    permutate prop, s
    permutate v, s
    pp = partition_points(prop)
    ret = makeM(rowSize(pp) - 1)
    Dim i As Long, j As Long, k As Long: k = 0
    For i = 0 To UBound(ret) Step 1
        ret(i) = makeM(pp(i + 1) - pp(i))
        For j = 0 To UBound(ret(i)) Step 1
            swapVariant ret(i)(j), v(k)
            k = k + 1
        Next j
    Next i
    v = Empty
    swapVariant group_by_pred, ret
End Function

'******************************************************************************
' csv�t�@�C����1�s��z��ɕ���
' expr      csv�t�@�C����1�s
' delim     ��؂蕶���i�ȗ����̓J���}�j
'******************************************************************************
Function csv2Vector(ByRef expr As Variant, Optional ByRef delimiter As Variant) As Variant
    Dim delim As String
    delim = IIf(VarType(delimiter) = vbString, delimiter, ",")
    Dim bn As Long, en As Long, counter As Long, isEven As Boolean
    Dim ret As Variant: ret = VBA.Array("")
    Dim LenExpr As Long: LenExpr = Len(expr)
    isEven = True
    bn = 1
    counter = -1
    Do
        For en = bn To LenExpr Step 1
            If Mid(expr, en, 1) = """" Then
                isEven = Not isEven
            ElseIf isEven And Mid(expr, en, 1) = delim Then
                counter = counter + 1
                ReDim Preserve ret(0 To counter)
                ret(counter) = Mid(expr, bn, en - bn)
                bn = en + 1
                Exit For
            End If
        Next en
        If bn < en Then
            counter = counter + 1
            ReDim Preserve ret(0 To counter)
            ret(counter) = Mid(expr, bn)
            bn = en + 1
        End If
     Loop While bn <= LenExpr
     Do While 0 <= counter
        If Left(ret(counter), 1) = """" Then
            ret(counter) = Mid(ret(counter), 2, Len(ret(counter)) - 2)
        End If
        ret(counter) = Replace(ret(counter), """""", """")
        ret(counter) = Replace(ret(counter), "\\t", vbLf)   ' \\t -> vbLf   Chr(10)
        ret(counter) = Replace(ret(counter), "\t", vbTab)    ' \t  -> vbTab  Chr(9)
        ret(counter) = Replace(ret(counter), vbLf, "\t")     ' vbLf   -> \\t
        counter = counter - 1
    Loop
    swapVariant csv2Vector, ret
End Function
    Public Function p_csv2Vector(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_csv2Vector = make_funPointer_with_2nd_Default(AddressOf csv2Vector, firstParam, secondParam)
    End Function

' �Q��1�����z��̋��L����(1)�Ɣ񋤗L����(0)�����
' �Ƃ��ɏ���comp�ŏ����\�[�g����Ă���O��
' �߂�l�͂��ꂼ��̔z��ɑ΂���󋵂��i�[����Jag�z��
Function A_overlap_B(ByRef a As Variant, ByRef b As Variant, Optional ByRef comp As Variant) As Variant
    If IsMissing(comp) Then
        A_overlap_B = A_overlap_B(a, b, p_less)
    Else
        Dim ppA As Variant, ppB As Variant
        Dim a2B As Variant, b2A As Variant
        ppA = partition_points_pred(a, comp)
        ppB = partition_points_pred(b, comp)
        a2B = mapF_swap(p_equal_range_pred(comp), b, subV(a, headN(ppA, -1)))
        a2B = mapF(p_less(p__n(0), p__n(1)), a2B)
        b2A = mapF_swap(p_equal_range_pred(comp), a, subV(b, headN(ppB, -1)))
        b2A = mapF(p_less(p__n(0), p__n(1)), b2A)
        A_overlap_B = VBA.Array( _
            foldl1(p_catV, zipWith(p_repeat, a2B, adjacent_op(p_minus(ph_2, ph_1), ppA))) _
            , _
            foldl1(p_catV, zipWith(p_repeat, b2A, adjacent_op(p_minus(ph_2, ph_1), ppB))) _
        )
    End If
    '   a = uniform_int_dist(20, 0, 20):  permutate a, sortIndex(a)  '  [ 0�`20]
    '   b = uniform_int_dist(20, 10, 30): permutate b, sortIndex(b)  '  [10�`30]
    '   x = A_overlap_B(a, b)
    ' -----------------------------------
    '   printM catR(a, x(0))
    '   0  0  1  4  5  5  5  6  9  10  10  13  13  15  15  15  16  17  17  19
    '   0  0  0  0  0  0  0  0  0   0   0   1   1   1   1   1   1   1   1   1
    ' -----------------------------------
    '   printM catR(b, x(1))
    '   12  13  13  14  15  16  16  17  18  19  21  22  22  23  23  24  26  26  28  30
    '    0   1   1   0   1   1   1   1   0   1   0   0   0   0   0   0   0   0   0   0
End Function
