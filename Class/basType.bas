Attribute VB_Name = "basType"
Option Explicit

Public Const SYSTEM_NAME = "�A�؎s�ꐸ�Z�V�X�e��"
Public Const PROGRAM_VERSION = "1.00"

'�t�H�[�J�X�Z�b�g���̐F
Public Const FOCUS_STOP_COLOR = &H80FF80
Public Const FOCUS_NO_COLOR = &HFFFFFF

'�����敪�{�^���N���b�N���̐F
Public Const BUTTON_ON = &H80C0FF
Public Const BUTTON_OFF = &H8000000F

'���W�X�g���L�[
Public Const REG_KEY = "Software\Com\Yawase"

'�[������
Public Const FRACTION_OMISSION = 1          '�؂�̂�
Public Const FRACTION_ROUNDOFF = 2          '�l�̌ܓ�
Public Const FRACTION_ROUNDUP = 3           '�؂�グ

'�ŋ敪
Public Const TAX_SOTOZEI = 1                '�O��
Public Const TAX_UCHIZEI = 2                '����
Public Const TAX_HIKAZEI = 3                '��ې�

'�̎����敪
Public Const RECEIPT_OFF = 0                '�o�͂��Ȃ�
Public Const RECEIPT_ON = 1                 '�o�͂���

'�Ɩ��敪
Public Const BUSINESS_DIV_EXHIBITION = 1    '�o�i
Public Const BUSINESS_DIV_BUYER = 2         '����
Public Const BUSINESS_DIV_ALL = 3           '����

'�m�F�\�o�͋敪
Public Const CHECK_REPORT_OFF = 0           '���o��
Public Const CHECK_REPORT_ON = 1            '�o�͍ς�

'�o�i�`�[�敪
Public Const EXHIBITION_REPORT_OFF = 0      '���o��
Public Const EXHIBITION_REPORT_ON = 1       '�o�͍ς�

'����`�[�敪
Public Const BUYER_REPORT_OFF = 0           '���o��
Public Const BUYER_REPORT_ON = 1            '�o�͍ς�

'�����敪
Public Const PAYMENT_OFF = 0                '������
Public Const PAYMENT_ON = 1                 '�����ς�

'�������
Public Const PAYMENT_DIV_CASH = 1           '����
Public Const PAYMENT_DIV_CHECK = 2          '���؎�
Public Const PAYMENT_DIV_TRANSFER = 3       '��s�U��

'�x���敪
Public Const SHIHARAI_OFF = 0               '������
Public Const SHIHARAI_ON = 1                '�x���ς�

'�x�����
Public Const SHIHARAI_DIV_CASH = 1          '����
Public Const SHIHARAI_DIV_CHECK = 2         '���؎�
Public Const SHIHARAI_DIV_TRANSFER = 3      '��s�U��

'���͍ς݋敪
Public Const INPUT_OFF = 0                  '������
Public Const INPUT_ON = 1                   '���͍ς�

'����`��
Public Const URIAGE_KEISIKI_KYOUBAI = 1     '����
Public Const URIAGE_KEISIKI_KOMUKAI = 2     '����

'�Ĕ��s�敪
Public Const REPRINT_OFF = 0                '���Ȃ�
Public Const REPRINT_ON = 1                 '����

'�����敪
Public Const HATIMONO_DIV_OFF = 0           '�����łȂ�
Public Const HATIMONO_DIV_ON = 1            '����

'�s���s�O�敪
Public Const TIKU_DIV_OFF = 0               '�s�O
Public Const TIKU_DIV_ON = 1                '�s��

'�����s�����敪
Public Const AUCTION_ON = 0                 '����
Public Const AUCTION_OFF = 1                '�s����

'�J�Ó�
Public Function HOLDING_DATE() As Variant
    '�P�E�W�E�P�T�E�Q�R
    HOLDING_DATE = Array("1", "8", "15", "23")
End Function

'���ʂ̊J�Ó�
Public Function MONTH_HOLDING_DATE(intMonth As Integer, intKaisu As Integer) As Variant

    Select Case intMonth
        Case 1:
            MONTH_HOLDING_DATE = Array("15", "23")
            intKaisu = 2
        Case 7:
            MONTH_HOLDING_DATE = Array("1", "8", "15")
            intKaisu = 3
        Case 8:
            MONTH_HOLDING_DATE = Null
            intKaisu = 0
        Case Else
            MONTH_HOLDING_DATE = Array("1", "8", "15", "23")
            intKaisu = 4
    End Select
    
End Function
