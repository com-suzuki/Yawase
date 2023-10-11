Attribute VB_Name = "basImectl"
'IME�̺�÷�Ă��擾���邽�߂�API
Private Declare Function ImmGetContext Lib "imm32.dll" _
        (ByVal hwnd As Long) As Long

'IME�̺�÷�Ă��J�����邽�߂�API
Private Declare Function ImmReleaseContext Lib "imm32.dll" _
        (ByVal hwnd As Long, ByVal himc As Long) As Long
        
'IME��On/Off��Ԃ�ݒ肷��API
Private Declare Function ImmSetOpenStatus Lib "imm32.dll" _
        (ByVal himc As Long, ByVal b As Long) As Long

'IME�̏����ϊ����������Ӱ�ނ�ݒ肷��API
Private Declare Function ImmSetConversionStatus Lib "imm32.dll" _
    (ByVal himc As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'IME�̏����ϊ����������Ӱ�ނ��擾����API
Private Declare Function ImmGetConversionStatus Lib "imm32.dll" _
        (ByVal himc As Long, lpdw As Long, lpdw2 As Long) As Long

'�쐬������÷�Ă���޳���֘A�t����API
Private Declare Function ImmAssociateContext Lib "imm32.dll" _
        (ByVal hwnd As Long, ByVal himc As Long) As Long

'����Ӱ�ނ̒萔
Private Const IME_CMODE_NATIVE = &H1    '���ړ���
Private Const IME_CMODE_KATAKANA = &H2  '����
Private Const IME_CMODE_LANGUAGE = &H3  '���{��
Private Const IME_CMODE_FULLSHAPE = &H8 '�S�p
Private Const IME_CMODE_ROMAN = &H10    '۰ώ�
Private Const IME_CMODE_CHARCODE = &H20 '���ޓ���

'// �ϊ����[�h
Public Enum IME_SMODE_ENUM
   IME_SMODE_NONE = &H0                   '// ���ϊ�
   IME_SMODE_PLAURALCLAUSE = &H1          '// �l���^�n��
   IME_SMODE_SINGLECONVERT = &H2          '// �V���O���L�����N�^���[�h(?)
   IME_SMODE_AUTOMATIC = &H4              '// ����
   IME_SMODE_PHRASEPREDICT = &H8          '// ���̃t���[�Y��\��(?)
   IME_SMODE_CONVERSATION = &H10          '// �b�����t�D��
End Enum

Dim lngTargetWindow As Long

'��  �I    �FIME�̓���Ӱ�ނ�ݒ�
'��  ��    �F
'��  ��    �F
'��  ��    �F
'               0:�Ȃ�
'               1:�I��
'               2:�I�t
'               3:�I�t�Œ�
'               4:�S�p�Ђ炪��
'               5:�S�p�J�^�J�i
'               6:���p����
'               7:�S�p�p��
'               8:���p�p��
'�߂�l    �F
'�쐬��    �F������� �R���E�G���W�j�A�����O ����
'�쐬�N�����F�P�X�X�W�^�R�^�P�R
'�X�V����  �F
'
Public Sub SetImeMode(lngNewTarget As Long, lngNewMode As Long)

    Dim lngAPIReVal As Long
    Dim lngInputMode As Long
    Dim lngConvertMode As Long
    Dim lngIMEHandle As Long

    On Error GoTo SetImeMode_Err

    'IME�̺�÷�Ă��J������
    lngAPIReVal = ImmReleaseContext(lngTargetWindow, lngIMEHandle)
    'IME�̺�÷�Ă��擾
    lngIMEHandle = ImmGetContext(lngNewTarget)
    '���݂̃R���g���[���̃n���h�����i�[
    lngTargetWindow = lngNewTarget
    
    'IME�̏����ϊ������Ɠ���Ӱ�ނ��擾
    lngAPIReVal = ImmGetConversionStatus(lngIMEHandle, lngInputMode, lngConvertMode)

    '����Ӱ�ނ̐ݒ�
    Select Case lngNewMode
        Case 0
            '�Ȃ�
        Case 1
            '�I��
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 1)
            Exit Sub
        Case 2
            '�I�t
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 0)
            Exit Sub
        Case 3
            '�I�t�Œ�
            
        Case 4
            '�S�p���[�h
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 1)
            '�S�p�Ђ炪��
            lngInputMode = lngInputMode And Not IME_CMODE_KATAKANA
            lngInputMode = lngInputMode Or IME_CMODE_FULLSHAPE Or IME_CMODE_NATIVE
            'lngConvertMode = IME_SMODE_ENUM.IME_SMODE_PHRASEPREDICT
        Case 5
            '�S�p���[�h
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 1)
            '�S�p�J�^�J�i�̏ꍇ
            lngInputMode = lngInputMode Or IME_CMODE_LANGUAGE Or _
                           IME_CMODE_FULLSHAPE Or IME_CMODE_KATAKANA
        Case 6
            '���p���[�h
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 0)
            '���p���ł̏ꍇ
            lngInputMode = lngInputMode And Not IME_CMODE_FULLSHAPE
            lngInputMode = lngInputMode Or IME_CMODE_LANGUAGE Or IME_CMODE_KATAKANA
        Case 7
            '���p���[�h
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 0)
            '�S�p�p���̏ꍇ
            lngInputMode = lngInputMode And Not IME_CMODE_LANGUAGE
            lngInputMode = lngInputMode Or IME_CMODE_FULLSHAPE
        Case 8
            '���p���[�h
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 0)
            '���p�p���̏ꍇ
            lngInputMode = lngInputMode And Not IME_CMODE_FULLSHAPE
            lngInputMode = lngInputMode And Not IME_CMODE_LANGUAGE
        Case 9
            '�S�p���[�h
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 1)
            '���p���ł̏ꍇ
            lngInputMode = lngInputMode And Not IME_CMODE_FULLSHAPE
            lngInputMode = lngInputMode Or IME_CMODE_LANGUAGE Or IME_CMODE_KATAKANA
    End Select

    'IME�̏����ϊ������Ɠ���Ӱ�ނ�ݒ�
    lngAPIReVal = ImmSetConversionStatus(lngIMEHandle, lngInputMode, lngConvertMode)

    Exit Sub

SetImeMode_Err:

    MsgBox Error$

End Sub
