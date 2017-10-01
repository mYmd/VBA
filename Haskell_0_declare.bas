Attribute VB_Name = "Haskell_0_declare"
'Haskell_0_declare
'Copyright (c) 2015 mmYYmmdd
Option Explicit

'======================================================
'          API�錾
'   Function Dimension          �z��̎����擾
'   Function placeholder        �v���[�X�z���_�E�I�u�W�F�N�g�̐���
'   Function is_placeholder     �v���[�X�z���_�E�I�u�W�F�N�g����
'   Function unbind_invoke      bind����Ă��Ȃ�VBA�֐���2�����ŌĂяo��
'   Function mapF_imple         �z��matrix�̊e�v�felem��Callback�֐���K�p����
'   Function zipWith            2�̔z��̊e�v�f�Ɋ֐���K�p����
'   Function foldl              �z��ɑ΂������̎��ɉ���������ݍ��݁i�����l�w�肠��j
'   Function foldr              �z��ɑ΂������̎��ɉ������E��ݍ��݁i�����l�w�肠��j
'   Function foldl1             �z��ɑ΂������̎��ɉ���������ݍ��݁i�擪�v�f�������l�Ƃ���j
'   Function foldr1             �z��ɑ΂������̎��ɉ������E��ݍ��݁i�����v�f�������l�Ƃ���j
'   Function scanl              �z��ɑ΂������̎��ɉ�������scan�i�����l�w�肠��j
'   Function scanr              �z��ɑ΂������̎��ɉ������Escan�i�����l�w�肠��j
'   Function scanl1             �z��ɑ΂������̎��ɉ�������scan�i�擪�v�f�������l�Ƃ���j
'   Function scanr1             �z��ɑ΂������̎��ɉ������Escan�i�����v�f�������l�Ƃ���j
'   Function stdsort            1�����z��̃\�[�g�C���f�b�N�X�o��
'   Function find_imple         �q��ɂ�錟��
'   Function find_best_imple    �q��ɂ��ŗǒl�ʒu����
'   Function repeat_imple       �֐��K�p�̃��[�v�i+ �I�������j
'   Function swapVariant        VARIANT�ϐ��ǂ����̃X���b�v
'   Sub      changeLBound       VBA�z���LBound�ύX
'   Function self_zipWith       1�����z��̗��ꂽ�v�f�Ԃ�2�������K�p����
'======================================================
' Callback�Ƃ��Ďg����֐��̃V�O�l�`����
' Function fun(ByRef x As Variant, ByRef y As Variant) As Variant
' ��������
' Function fun(ByRef x As Variant, Optional ByRef dummy As Variant) As Variant
'======================================================
' VBA�z��̎����擾
Declare PtrSafe Function Dimension Lib "mapM.dll" (ByRef v As Variant) As Long

'�v���[�X�z���_�E�I�u�W�F�N�g�̐���
Declare PtrSafe Function placeholder Lib "mapM.dll" (Optional ByVal N As Long = 0) As Variant

'�v���[�X�z���_�E�I�u�W�F�N�g����
Declare PtrSafe Function is_placeholder Lib "mapM.dll" (ByRef v As Variant) As Long

'bind����Ă��Ȃ�VBA�֐���2�����ŌĂяo��
Declare PtrSafe Function unbind_invoke Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef param1 As Variant, _
        ByRef param2 As Variant) As Variant

' �z��matrix�̊e�v�felem��Callback�֐���K�p����
Declare PtrSafe Function mapF_imple Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant) As Variant

'�z��matrix1��matrix2�̊e�v�f��2�ϐ���Callback�֐���K�p����
Declare PtrSafe Function zipWith Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix1 As Variant, _
        ByRef matrix2 As Variant) As Variant

' 3�����܂ł�VBA�z��ɑ΂������̎��ɉ���������ݍ��݁i�����l�w�肠��j
Declare PtrSafe Function foldl Lib "mapM.dll" ( _
                    ByRef pCallback As Variant, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3�����܂ł�VBA�z��ɑ΂������̎��ɉ������E��ݍ��݁i�����l�w�肠��j
Declare PtrSafe Function foldr Lib "mapM.dll" ( _
                    ByRef pCallback As Variant, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3�����܂ł�VBA�z��ɑ΂������̎��ɉ���������ݍ��݁i�擪�v�f�������l�Ƃ���j
Declare PtrSafe Function foldl1 Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3�����܂ł�VBA�z��ɑ΂������̎��ɉ������E��ݍ��݁i�����v�f�������l�Ƃ���j
Declare PtrSafe Function foldr1 Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3�����܂ł�VBA�z��ɑ΂������̎��ɉ�������scan�i�����l�w�肠��j
Declare PtrSafe Function scanl Lib "mapM.dll" ( _
                    ByRef pCallback As Variant, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3�����܂ł�VBA�z��ɑ΂������̎��ɉ������Escan�i�����l�w�肠��j
Declare PtrSafe Function scanr Lib "mapM.dll" ( _
                    ByRef pCallback As Variant, _
                ByRef init As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3�����܂ł�VBA�z��ɑ΂������̎��ɉ�������scan�i�擪�v�f�������l�Ƃ���j
Declare PtrSafe Function scanl1 Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 3�����܂ł�VBA�z��ɑ΂������̎��ɉ������Escan�i�����v�f�������l�Ƃ���j
Declare PtrSafe Function scanr1 Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        Optional ByVal axis As Long = 1) As Variant

' 1�����z��̃\�[�g�C���f�b�N�X�o��
Declare PtrSafe Function stdsort Lib "mapM.dll" (ByRef ary As Variant, _
                                         ByVal defaultFlag As Long, _
                                         ByRef pComp As Variant) As Variant

' �q��ɂ��ʒu����
Declare PtrSafe Function find_imple Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        ByVal def As Long) As Long

' �q��ɂ��ŗǒl�ʒu����
Declare PtrSafe Function find_best_imple Lib "mapM.dll" ( _
                ByRef pCallback As Variant, _
            ByRef matrix As Variant, _
        ByVal def As Long) As Long

'�֐��K�p�̃��[�v�i+ �I�������j
Declare PtrSafe Function repeat_imple Lib "mapM.dll" ( _
                        ByRef init As Variant, _
                    ByRef pred As Variant, _
                ByRef trans As Variant, _
            ByVal maxN As Long, _
        ByVal scan As Long, _
    ByVal stopCondition As Long) As Variant

'VARIANT�ϐ��ǂ����̃X���b�v
Declare PtrSafe Function swapVariant Lib "mapM.dll" (ByRef a As Variant, ByRef b As Variant) As Long

' VBA�z���LBound�ύX
Declare PtrSafe Sub changeLBound Lib "mapM.dll" (ByRef v As Variant, ByVal lbound_v As Long)

' 1�����z��̗��ꂽ�v�f�Ԃ�2�������K�p����
Declare PtrSafe Function self_zipWith Lib "mapM.dll" ( _
                                ByRef pCallback As Variant, _
                            ByRef vec As Variant, _
                      ByVal shift As Long) As Variant
