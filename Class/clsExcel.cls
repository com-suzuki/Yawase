VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExcelOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_varTitleName As Variant
Private m_AdoRecordset As Object

'Windows Api
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'OPENFILENAME�\����
Private Type OPENFILENAME
    lStructSize As Long             '�\���̂̃T�C�Y
    hwndOwner As Long               '�E�C���h�E�̃n���h��
    hInstance As Long               '�C���X�^���X�n���h��
    lpstrFilter As String           '�t�B���^�[
    lpstrCustomFilter As String     '�J�X�^���t�B���^�[
    nMaxCustFilter As Long          '�J�X�^���t�B���^�[�̃T�C�Y
    nFilterIndex As Long            '�t�B���^�̃C���f�b�N�X
    lpstrFile As String             '�t�@�C�����̃o�b�t�@
    nMaxFile As Long                '�t�@�C�����̃o�b�t�@�̃T�C�Y
    lpstrFileTitle As String        '�t���p�X�p�̃o�b�t�@
    nMaxFileTitle As Long           '�t���p�X�p�̃o�b�t�@�̃T�C�Y
    lpstrInitialDir As String       '�f�B���N�g�����w��
    lpstrTitle As String            '�_�C�����O�{�b�N�X�̃^�C�g��
    flags As Long                   '�萔�iOFN�Q�~�~�Q�Ɓj
    nFileOffset As Integer          '�t���p�X�̒��̃t�@�C�����܂ł̃I�t�Z�b�g
    nFileExtension As Integer       '�g���q�܂ł̃I�t�Z�b�g
    lpstrDefExt As String           '�f�t�H���g�̊g���q
    lCustData As Long               'lpfnHook�œn���f�[�^
    lpfnHook As Long                '�t�b�N�֐��̃|�C���^
    lpTemplateName As String        '�e���v���[�g��
End Type

Private Const OFN_ALLOWMULTISELECT = &H200       '�����t�@�C����I���\�ɂ���
Private Const OFN_CREATEPROMPT = &H2000          '�w��̃t�@�C���������݂��Ȃ����Ƀ��b�Z�[�W�{�b�N�X��\��
Private Const OFN_FILEMUSTEXIST = &H1000         '���݂��Ȃ��t�@�C�����͓��͕s��
Private Const OFN_HIDEREADONLY = &H4             '�ǂݎ���p�̃`�F�b�N�{�b�N�X���\��
Private Const OFN_NOCHANGEDIR = &H8              '���̃T�u�f�B���N�g������I��s��
Private Const OFN_NOREADONLYRETURN = &H8000      '�ǂݍ��ݐ�p�t�@�C���Ə������݋֎~�f�B���N�g���̑I��s��
Private Const OFN_NOVALIDATE = &H100             '�t�@�C�����̗L�������`�F�b�N���Ȃ�
Private Const OFN_OVERWRITEPROMPT = &H2          '�����̃t�@�C�������w�肵�����Ƀ��b�Z�[�W���o��
Private Const OFN_PATHMUSTEXIST = &H800          '�L���ȃp�X��������������
Private Const OFN_READONLY = &H1                 '�ǂݎ���p�̃`�F�b�N�{�b�N�X���`�F�b�N
Private Const OFN_SHOWHELP = &H10                '�w���v�{�^����\��

'�ځ@�I�@�@�FExcel�o�͏���(Excel�̃o�[�W�����Ɉˑ����Ȃ�)
'���@���@�@�F�Q�Ɛݒ��Microsoft Excel8.0 Object Library���Q�Ƃ��Ȃ�
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function OutPut() As Boolean

    Dim objExcelApp As Object       'Excel Application�I�u�W�F�N�g
    Dim objExcelBook As Object      'Excel Book �I�u�W�F�N�g
    Dim objExcelSheets As Object    'Excel Sheets �I�u�W�F�N�g
    Dim objExcelSheet As Object     'Excel Sheet �I�u�W�F�N�g
    Dim strFilter As String
    Dim strSaveFile As String
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo OutPut_Err

    OutPut = False

    '�o�͐�w��
    strFilter = "Excel�t�@�C��(*.xls)" & Chr(0) & "*.xls" & Chr(0)
    strFilter = strFilter & "���ׂẴt�@�C��(*.*)" & Chr(0) & "*.*" & Chr(0)
    strSaveFile = OpenSaveDlg(0, strFilter, "", "xls")
    If Trim(strSaveFile) = "" Then
        Exit Function
    End If
    
    '�}�E�X�|�C���^�������v�ɕύX
    Screen.MousePointer = vbHourglass

    'Excel�I�u�W�F�N�g�擾
    Set objExcelApp = CreateObject("Excel.Application")
    
    '�V�[�g�̐V�K�쐬
    objExcelApp.Workbooks.Add
    Set objExcelBook = objExcelApp.ActiveWorkbook
    Set objExcelSheets = objExcelBook.Worksheets
    Set objExcelSheet = objExcelBook.Sheets(1)
    objExcelApp.Visible = True
    
    '�^�C�g���̏o��
    lngRow = 1
    If Not IsNull(m_varTitleName) Then
        For lngCol = 0 To UBound(m_varTitleName)
            objExcelSheet.Cells(lngRow, lngCol + 1).Value = m_varTitleName(lngCol)
        Next lngCol
        lngRow = lngRow + 1
    End If
    
    '�f�[�^�o��
    If IsObject(m_AdoRecordset) = True Then
        Do While Not m_AdoRecordset.EOF
            For lngCol = 0 To m_AdoRecordset.Fields.Count - 1
                If Not IsNull(m_AdoRecordset.Fields(lngCol).Value) Then
                    objExcelSheet.Cells(lngRow, lngCol + 1).Value = m_AdoRecordset.Fields(lngCol).Value
                End If
            Next lngCol
            m_AdoRecordset.MoveNext
            lngRow = lngRow + 1
        Loop
    End If
    
    '�V�[�g��ۑ�
    objExcelSheet.SaveAs strSaveFile
    
    '�}�E�X�|�C���^�����ɖ߂�
    Screen.MousePointer = vbDefault

    OutPut = True

    Exit Function

OutPut_Err:

    '�}�E�X�|�C���^�����ɖ߂�
    Screen.MousePointer = vbDefault
    Set objExcelSheets = Nothing
    OutPut = False
    Call MsgBox("Excel�o�͏����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "OutPut_Err")
    
End Function

Private Function OpenSaveDlg(lngHandle As Long, strFilter As String, strInitialDir As String, strDefExt As String) As String

    Dim OFN As OPENFILENAME                     'OPENFILENAME�\����
    Dim lngRet As Long                          '�߂�l

    On Error GoTo OpenSaveDlg_Err

    With OFN                                    '�\���̂̐ݒ�
        .flags = OFN_PATHMUSTEXIST Or _
                 OFN_FILEMUSTEXIST Or _
                 OFN_HIDEREADONLY Or _
                 OFN_OVERWRITEPROMPT
                 'Or OFN_SHOWHELP
        .hInstance = App.hInstance              '�C���X�^���X�n���h����ݒ�
        .hwndOwner = lngHandle                  '�E�C���h�E�n���h����ݒ�
        '.lpstrTitle = "�t�@�C����ۑ�����"      '�R�����_�C�A���O�̃^�C�g����ݒ�
        .lpstrFilter = strFilter                '�t�B���^�[��ݒ�
        .lStructSize = Len(OFN)                 '�\���̂̃T�C�Y��ݒ�
        .nMaxFile = 250                         '�t�@�C�����̃o�b�t�@�T�C�Y��ݒ�
        .lpstrFileTitle = String(250, Chr(0))   '�t���p�X�p�̃o�b�t�@���m��
        .nMaxFileTitle = 250                    '�t���p�X�p�̃o�b�t�@�T�C�Y��ݒ�
        .lpstrFile = String(250, Chr(0))        '�t�@�C�����̃o�b�t�@���m��
        .lpstrInitialDir = strInitialDir        '�f�t�H���g�̃f�B���N�g����ݒ�
        .lpstrDefExt = strDefExt                '�f�t�H���g�̊g���q��ݒ�
    End With
    
    lngRet = GetSaveFileName(OFN)
    
    If lngRet = 0 Then
         OpenSaveDlg = vbNullString
    Else
         OpenSaveDlg = OFN.lpstrFile             '�p�X���擾
    End If

    Exit Function

OpenSaveDlg_Err:

    OpenSaveDlg = ""
    Call MsgBox("�ۑ��_�C�A���O�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "OpenSaveDlg_Err")

End Function

Public Property Get TitleName() As Variant

    TitleName = m_varTitleName

End Property

Public Property Let TitleName(ByVal Param As Variant)

    m_varTitleName = Param

End Property

Public Property Get RecordSet() As Object

    Set RecordSet = m_AdoRecordset

End Property

Public Property Let RecordSet(ByVal Param As Object)

    Set m_AdoRecordset = Param

End Property

Private Sub Class_Initialize()

    m_varTitleName = Null

End Sub