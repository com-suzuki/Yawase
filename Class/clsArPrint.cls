VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************
'�N���X���@�FActiveReport1.5 �p�N���X(ADO�o�[�W����)
'
'���@���@�@�F�P�D�Q�Ɛݒ�  ActiveReports Viewer Control 1.5J
'�@�@�@�@�@�@�@�@�@�@�@�@�@Data Dynamics ActiveReports 1.5J
'�@�@�@�@�@�@�@�@�@�@�@�@�@ActiveReports PDF Export Fileter
'�@�@�@�@�@�@�@�@�@�@�@�@�@ActiveReports PDF Export Fileter Fixer
'
'           �Q�D�f�[�^�R���g���[������DataControl1�ɂ��邱��
'
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�o�[�W�����F1.00
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'********************************************************************************

Private Const INI_FILENAME = "ArPrint.ini"
Private m_strIniFileName As String

Private Type PrintInfo
    Name                    As String           '���|�[�g����
    'objReport               As Object           'ActiveReport�I�u�W�F�N�g
    objReport               As ActiveReport     'ActiveReport�I�u�W�F�N�g
    DatabaseName            As String           '�f�[�^�x�[�X��
    Connection              As String           '�f�[�^�x�[�X�R�l�N�V����
    SQL                     As String           'SQL��
    Caption                 As String           '�v���r���[�E�C���h�E�̃L���v�V����
    TableForExcel           As String           'Excel�p���[�N�e�[�u����
    PrnDefaultPrinter       As Boolean          'True:�ʏ�g���v�����^
    PrnDeviceName           As String           '�v�����^��
    PrnPaperSize            As Integer          '�p���T�C�Y
    PrnOrientation          As Integer          '�p���̌���
    PrnPageTopMargin        As Integer          '�]���i��j
    PrnPageBottomMargin     As Integer          '�]���i���j
    PrnPageLeftMargin       As Integer          '�]���i���j
    PrnPageRightMargin      As Integer          '�]���i�E�j
    NoData                  As Boolean          '�f�[�^�L��
    PrnPaperSource          As Integer          '�������@
End Type
Private m_PrintInfo As PrintInfo

'INI�t�@�C����API�錾
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

'[̧�ق��J��]��[̧�ق�ۑ�]�޲�۸��ޯ���ׂ̈̍\���̐錾
Private Type OPENFILENAME
    lStructSize As Long         '�\���̻̂���
    hwndOwner As Long           '�޲�۸��ޯ�������L���鳨��޳�������
    hInstance As Long           '���ع���� �ݽ�ݽ
    lpstrFilter As String       '̨���
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long        '��̫�Ă�̨��
    lpstrFile As String         '�I�����ꂽ̧�ٖ�
    nMaxFile As Long            '̧�ٖ��̍ő咷
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long               '��߼��
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'[̧�ق̕ۑ�]�޲�۸��ޯ�����Ăяo��API
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'OPENFILENAME�\���̂�flags�Ŏw�肷��萔
Private Const OFN_READONLY = &H1                '[�ǂݎ���p]�����ޯ��������
Private Const OFN_OVERWRITEPROMPT = &H2         '̧�ق����݂��Ă����ꍇ�A�㏑����₢���킹��
Private Const OFN_HIDEREADONLY = &H4            '[�ǂݎ���p]�����ޯ�����\��
Private Const OFN_SHOWHELP = &H10               '[����]���݂̕\��
Private Const OFN_ALLOWMULTISELECT = &H200      '������̧�ق�I���\��
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800         '���݂��Ȃ��߽������͕s��
Private Const OFN_FILEMUSTEXIST = &H1000        '���݂��Ȃ�̧�ٖ�����͕s��
Private Const OFN_CREATEPROMPT = &H2000         '̧�ق����݂��Ȃ������ꍇ�A�V�K�쐬���邩�ǂ����\��
Private Const OFN_EXPLORER = &H80000

'�ځ@�I�@�@�F�N���X������
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub Class_Initialize()

    On Error GoTo Class_Initialize_Err

    'INI�t�@�C�����ݒ�
    If right(Trim(App.Path), 1) = "\" Then
        m_strIniFileName = App.Path & INI_FILENAME
    Else
        m_strIniFileName = App.Path & "\" & INI_FILENAME
    End If
    
    '�v���p�e�B�̏�����
    Set m_PrintInfo.objReport = Nothing
    
    '�W���̃v���p�e�B�l��ݒ�
    With m_PrintInfo
        .Name = ""
        .DatabaseName = ""
        .Connection = ""
        .SQL = ""
        .Caption = ""
        .TableForExcel = ""
        .PrnDefaultPrinter = True
        .PrnDeviceName = ""
        .PrnPaperSize = vbPRPSA4
        .PrnOrientation = vbPRORLandscape
        .PrnPageTopMargin = 10
        .PrnPageBottomMargin = 10
        .PrnPageLeftMargin = 10
        .PrnPageRightMargin = 10
        .PrnPaperSource = 0
        .NoData = False
    End With
    
    Exit Sub
    
Class_Initialize_Err:
    
   Call MsgBox("�N���X�������G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Class_Initialize_Err")
    
End Sub

'�ځ@�I�@�@�F�N���X�I����
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub Class_Terminate()
    
    '�����Ȃ�

End Sub

'�ځ@�I�@�@�FActiveReport�̏o�͏���
'���@���@�@�F�f�[�^�R���g���[������DataControl1�ɂ��邱��
'���@�ʁ@�@�F
'���@���@�@�F0:����v���r���[ 1:��� 2:PDF�o��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function PrintActiveReport(intDestination As Integer, Optional strPdfFileName As String) As Boolean
    
    On Error GoTo PrintActiveReport_Err
    
    PrintActiveReport = False
    
    '�v���p�e�B�X�V
    If Refresh() = False Then Exit Function
    
    Select Case intDestination
        Case 0:     '����v���r���[
            '�}�E�X�|�C���^�������v�ɕύX
            Screen.MousePointer = vbHourglass
        
            '�V�������|�[�g��ARViewer�R���g���[���Ɋ��蓖�Ă܂�
            Set frmARPreview.arv.ReportSource = m_PrintInfo.objReport
            frmARPreview.Caption = m_PrintInfo.Caption
            Set frmARPreview.m_objArPrint = Me
            m_PrintInfo.objReport.Run
            DoEvents
            
            '�}�E�X�|�C���^�����ɖ߂�
            Screen.MousePointer = vbDefault
            If m_PrintInfo.NoData = False Then
                frmARPreview.Show vbModal
            End If
        Case 1:     '����_�C�A���O��\�������Ɉ��
            m_PrintInfo.objReport.PrintReport False
        Case 2:     'PDF�t�@�C���֏o��
            '�o�͐�w��
            If IsMissing(strPdfFileName) Or strPdfFileName = "" Then
                strPdfFileName = Get_FileName(0)
                If Trim(strPdfFileName) = "" Then
                    Exit Function
                End If
            End If
        
            '�}�E�X�|�C���^�������v�ɕύX
            Screen.MousePointer = vbHourglass
        
            Dim objARExport As New ARExportPDF
            Dim objARExportFix As New ARExportPDFFix
            objARExport.FileName = strPdfFileName
            m_PrintInfo.objReport.Run
            m_PrintInfo.objReport.Export objARExport
            'PDF�t�@�C���̕����������C��
            objARExportFix.ExportFix objARExport.FileName
    
            '�}�E�X�|�C���^�����ɖ߂�
            Screen.MousePointer = vbDefault
    End Select

    PrintActiveReport = True

    Exit Function

PrintActiveReport_Err:

    PrintActiveReport = False
    '�}�E�X�|�C���^�����ɖ߂�
    Screen.MousePointer = vbDefault
    Call MsgBox("ActiveReport�̏o�͏����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "PrintActiveReport_Err")

End Function

'�ځ@�I�@�@�F�v���p�e�B�X�V
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Refresh() As Boolean
    
    On Error GoTo Refresh_Err
    
    Screen.MousePointer = vbHourglass
    
    Refresh = False
    
    Call SetIni(m_PrintInfo.Name, "DefaultPrinter", m_PrintInfo.PrnDefaultPrinter)
    If m_PrintInfo.PrnDefaultPrinter = True Then
        m_PrintInfo.PrnDeviceName = ""
    End If
    Call SetIni(m_PrintInfo.Name, "DeviceName", m_PrintInfo.PrnDeviceName)
    Call SetIni(m_PrintInfo.Name, "PaperSize", m_PrintInfo.PrnPaperSize)
    Call SetIni(m_PrintInfo.Name, "Orientation", m_PrintInfo.PrnOrientation)
    Call SetIni(m_PrintInfo.Name, "PageTopMargin", m_PrintInfo.PrnPageTopMargin)
    Call SetIni(m_PrintInfo.Name, "PageBottomMargin", m_PrintInfo.PrnPageBottomMargin)
    Call SetIni(m_PrintInfo.Name, "PageLeftMargin", m_PrintInfo.PrnPageLeftMargin)
    Call SetIni(m_PrintInfo.Name, "PageRightMargin", m_PrintInfo.PrnPageRightMargin)
    Call SetIni(m_PrintInfo.Name, "PrnPaperSource", m_PrintInfo.PrnPaperSource)
    
    With m_PrintInfo.objReport
        '�f�[�^�x�[�X�ݒ�
        .DataControl1.ConnectionString = m_PrintInfo.Connection
        DoEvents
        If Trim(m_PrintInfo.SQL) <> "" Then
            .DataControl1.Source = m_PrintInfo.SQL
        End If
        .DataControl1.Refresh
        DoEvents
        '�v�����^�ݒ�
        On Error GoTo Printer_Setup
        If Trim(m_PrintInfo.PrnDeviceName) <> "" Then .Printer.DeviceName = m_PrintInfo.PrnDeviceName
        On Error GoTo Refresh_Err
        .Printer.PaperSize = m_PrintInfo.PrnPaperSize
        .Printer.Orientation = m_PrintInfo.PrnOrientation
        If m_PrintInfo.PrnPaperSource <> 0 Then
            .Printer.PaperBin = m_PrintInfo.PrnPaperSource
        End If
        '�]���ݒ�
        .PageTopMargin = mm_To_Twips(m_PrintInfo.PrnPageTopMargin)
        .PageBottomMargin = mm_To_Twips(m_PrintInfo.PrnPageBottomMargin)
        .PageLeftMargin = mm_To_Twips(m_PrintInfo.PrnPageLeftMargin)
        .PageRightMargin = mm_To_Twips(m_PrintInfo.PrnPageRightMargin)
        '�h�L�������g��
        .documentName = m_PrintInfo.Caption
    End With

    Screen.MousePointer = vbDefault

    Refresh = True

    Exit Function

Printer_Setup:

    Screen.MousePointer = vbDefault
    Call MsgBox("�����ȃv�����^���ݒ肳��Ă��܂��B" & vbCrLf & "�v�����^�̐ݒ���s���Ă��������B", vbOKOnly + vbCritical, "�G���[")
    
    '�y�[�W�ݒ�_�C�A���O�̕\��
    m_PrintInfo.PrnDeviceName = ""
    Set frmARPage.m_objArPrint = Me
    frmARPage.Show vbModal
    Refresh = Not frmARPage.m_blnCanselFlg
    
    Exit Function

Refresh_Err:

    Refresh = False
    Screen.MousePointer = vbDefault
    Call MsgBox("Refresh�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Refresh_Err")

End Function

'�ځ@�I�@�@�F�~������Twips�ւ̕ϊ�
'���@���@�@�F�����֐�
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Function mm_To_Twips(intmm As Integer) As Integer
    
    Dim intBuff1 As Integer
    
    On Error Resume Next
    
    '��1������567Twips 1mm��56.7Twips
    intBuff1 = intmm * 56.7
    mm_To_Twips = Fix(intBuff1)

End Function

'�ځ@�I�@�@�F�ݒ�ǂݍ���
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Function ReadPrnIni(strName As String) As Boolean
    
    On Error GoTo ReadPrnIni_Err
    
    If GetIni(strName, "DefaultPrinter") <> "" Then
        If GetIni(strName, "DefaultPrinter") = "True" Then
            m_PrintInfo.PrnDefaultPrinter = True
        ElseIf GetIni(strName, "DefaultPrinter") = "False" Then
            m_PrintInfo.PrnDefaultPrinter = False
        Else
            m_PrintInfo.PrnDefaultPrinter = True
        End If
    Else
        m_PrintInfo.PrnDefaultPrinter = True
    End If
    If m_PrintInfo.PrnDefaultPrinter = True Then
        m_PrintInfo.PrnDeviceName = Printer.DeviceName
    Else
        m_PrintInfo.PrnDeviceName = GetIni(strName, "DeviceName")
    End If
    If GetIni(strName, "PaperSize") <> "" Then
        m_PrintInfo.PrnPaperSize = GetIni(strName, "PaperSize")
    End If
    If GetIni(strName, "Orientation") <> "" Then
        m_PrintInfo.PrnOrientation = GetIni(strName, "Orientation")
    End If
    If GetIni(strName, "PageTopMargin") <> "" Then
        m_PrintInfo.PrnPageTopMargin = GetIni(strName, "PageTopMargin")
    End If
    If GetIni(strName, "PageBottomMargin") <> "" Then
        m_PrintInfo.PrnPageBottomMargin = GetIni(strName, "PageBottomMargin")
    End If
    If GetIni(strName, "PageLeftMargin") <> "" Then
        m_PrintInfo.PrnPageLeftMargin = GetIni(strName, "PageLeftMargin")
    End If
    If GetIni(strName, "PageRightMargin") <> "" Then
        m_PrintInfo.PrnPageRightMargin = GetIni(strName, "PageRightMargin")
    End If
    If GetIni(strName, "PrnPaperSource") <> "" Then
        m_PrintInfo.PrnPaperSource = GetIni(strName, "PrnPaperSource")
    End If
    
    ReadPrnIni = True

    Exit Function

ReadPrnIni_Err:

    ReadPrnIni = False
    Call MsgBox("�ݒ�ǂݍ��݃G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ReadPrnIni_Err")

End Function

'�ځ@�I�@�@�FINI�t�@�C������̃f�[�^�擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Function GetIni(ByVal vstrKeyword As String, ByVal vstrIniFile As String) As String
    
    Dim strResult   As String * 1024
    Dim intTemp     As Integer
    Dim intLen      As Integer
    Dim strFileName As String
   
    On Error Resume Next
   
    intTemp = GetPrivateProfileString(vstrKeyword, vstrIniFile, "", strResult, Len(strResult), m_strIniFileName)
    GetIni = left$(strResult, intTemp)

End Function

'�ځ@�I�@�@�FINI�t�@�C���փf�[�^�X�V
'���@���@�@�F�����֐�
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Function SetIni(ByVal strSection As String, ByVal vstrKeyword As String, ByVal vstrKeyVal As String) As Integer
    
    Dim intTemp As Integer

    On Error Resume Next

    intTemp = WritePrivateProfileString(strSection, vstrKeyword, vstrKeyVal, m_strIniFileName)
    SetIni = intTemp

End Function

'�ځ@�I�@�@�F�ۑ��_�C�A���O��\�����ăt�@�C�������擾����
'���@���@�@�F�����֐�
'���@�ʁ@�@�F
'���@���@�@�F0:PDF�t�@�C�� 1:Excel�t�@�C��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Function Get_FileName(intFlg As Integer) As String
    
    Dim typOpenFileName As OPENFILENAME
    Dim lngHWnd As Long
    Dim strFilter As String
    Dim strDefDir As String
    
    On Error GoTo Get_FileName_Err
    
    If intFlg = 0 Then
        strFilter = "PDF̧��(*.PDF)" & vbNullChar & "*.pdf" & vbNullChar & _
                    "�S�Ă�̧��(*.*)" & vbNullChar & "*.*" & vbNullChar
    ElseIf intFlg = 1 Then
        strFilter = "Exceļ��(*.XLS)" & vbNullChar & "*.xls" & vbNullChar & _
                    "�S�Ă�̧��(*.*)" & vbNullChar & "*.*" & vbNullChar
    End If
    lngHWnd = 0
    strDefDir = ""
    
    With typOpenFileName
        '���ނ̐ݒ�
        .lStructSize = Len(typOpenFileName)
        '�eWindow�̎w��
        .hwndOwner = lngHWnd
        '���ع���݂̲ݽ�ݽ���w��
        .hInstance = App.hInstance
        '̨���̎�ނ�ݒ�
        .lpstrFilter = strFilter
        '��è�ނ�̨���̔ԍ���ݒ�
        .nFilterIndex = 1
        '[̧�ٖ�]�ޯ���̓��e��������
        .lpstrFile = String(256, Chr(0))
        '�ő�̧�ْ��̐ݒ�
        .nMaxFile = 256
        '̧�ق����ق��󂯎���߲��
        .lpstrFileTitle = String(256, Chr(0))
        '�ő�̧�ْ��̐ݒ�
        .nMaxFileTitle = 256
        '��̫���ިڸ�؂̐ݒ�
        .lpstrInitialDir = strDefDir
        '�޲�۸ނ����ق�ݒ�
        '.lpstrTitle = "̧�ق��J��"
        '��߼�݂̐ݒ�
        .Flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST _
            Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    End With
    
    '[̧�ق��J��]�޲�۸ނ̕\��
    If GetSaveFileName(typOpenFileName) = 0 Then
        '��ݾق��ꂽ�ꍇ
        Get_FileName = ""
    Else
        'OK���݂������ꂽ�ꍇ�A̧�ٖ��̕\��
        'Null�����̏���
        If InStr(typOpenFileName.lpstrFile, ".") > 0 Then
            Get_FileName = left(typOpenFileName.lpstrFile, _
                    InStr(typOpenFileName.lpstrFile, vbNullChar) - 1)
        Else
            '�g���q���Ȃ��ꍇ
            If intFlg = 0 Then
                Get_FileName = left(typOpenFileName.lpstrFile, _
                        InStr(typOpenFileName.lpstrFile, vbNullChar) - 1) & _
                        ".PDF"
            ElseIf intFlg = 1 Then
                Get_FileName = left(typOpenFileName.lpstrFile, _
                        InStr(typOpenFileName.lpstrFile, vbNullChar) - 1) & _
                        ".XLS"
            End If
        End If
    End If

    Exit Function

Get_FileName_Err:

    Get_FileName = ""
    Call MsgBox("�t�@�C�����擾�����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Get_FileName_Err")

End Function

Public Property Get Name() As String

    Name = m_PrintInfo.Name
    
End Property

Public Property Let Name(Param As String)

    m_PrintInfo.Name = Param
    '�ݒ�̓ǂݍ���
    Call ReadPrnIni(m_PrintInfo.Name)
    
End Property

Public Property Get objReport() As Object

    Set objReport = m_PrintInfo.objReport
    
End Property

Public Property Let objReport(Param As Object)

    Set m_PrintInfo.objReport = Param
    
End Property

Public Property Get DatabaseName() As String

    DatabaseName = m_PrintInfo.DatabaseName
    
End Property

Public Property Let DatabaseName(Param As String)

    m_PrintInfo.DatabaseName = Param
    
End Property

Public Property Get Caption() As String

    Caption = m_PrintInfo.Caption
    
End Property

Public Property Let Caption(Param As String)

    m_PrintInfo.Caption = Param
    
End Property

Public Property Get TableForExcel() As String

    TableForExcel = m_PrintInfo.TableForExcel
    
End Property

Public Property Let TableForExcel(Param As String)

    m_PrintInfo.TableForExcel = Param
    
End Property

Public Property Get PrnDefaultPrinter() As Boolean

    PrnDefaultPrinter = m_PrintInfo.PrnDefaultPrinter
    
End Property

Public Property Let PrnDefaultPrinter(Param As Boolean)

    m_PrintInfo.PrnDefaultPrinter = Param
    
End Property

Public Property Get PrnDeviceName() As String

    PrnDeviceName = m_PrintInfo.PrnDeviceName
    
End Property

Public Property Let PrnDeviceName(Param As String)

    m_PrintInfo.PrnDeviceName = Param
    
End Property

Public Property Get PrnPaperSize() As String

    PrnPaperSize = m_PrintInfo.PrnPaperSize
    
End Property

Public Property Let PrnPaperSize(Param As String)

    m_PrintInfo.PrnPaperSize = Param
    
End Property

Public Property Get PrnOrientation() As String

    PrnOrientation = m_PrintInfo.PrnOrientation
    
End Property

Public Property Let PrnOrientation(Param As String)

    m_PrintInfo.PrnOrientation = Param
    
End Property

Public Property Get PrnPageTopMargin() As String

    PrnPageTopMargin = m_PrintInfo.PrnPageTopMargin
    
End Property

Public Property Let PrnPageTopMargin(Param As String)

    m_PrintInfo.PrnPageTopMargin = Param
    
End Property

Public Property Get PrnPageBottomMargin() As String

    PrnPageBottomMargin = m_PrintInfo.PrnPageBottomMargin
    
End Property

Public Property Let PrnPageBottomMargin(Param As String)

    m_PrintInfo.PrnPageBottomMargin = Param
    
End Property

Public Property Get PrnPageLeftMargin() As String

    PrnPageLeftMargin = m_PrintInfo.PrnPageLeftMargin
    
End Property

Public Property Let PrnPageLeftMargin(Param As String)

    m_PrintInfo.PrnPageLeftMargin = Param
    
End Property

Public Property Get PrnPageRightMargin() As String

    PrnPageRightMargin = m_PrintInfo.PrnPageRightMargin
    
End Property

Public Property Let PrnPageRightMargin(Param As String)

    m_PrintInfo.PrnPageRightMargin = Param
    
End Property

Public Property Get PrnPaperSource() As Integer

    PrnPaperSource = m_PrintInfo.PrnPaperSource
    
End Property

Public Property Let PrnPaperSource(Param As Integer)

    m_PrintInfo.PrnPaperSource = Param
    
End Property

Public Property Get Connection() As String

    Connection = m_PrintInfo.Connection
    
End Property

Public Property Let Connection(Param As String)

    m_PrintInfo.Connection = Param
    
End Property

Public Property Get SQL() As String

    SQL = m_PrintInfo.SQL
    
End Property

Public Property Let SQL(Param As String)

    m_PrintInfo.SQL = Param
    
End Property

Public Property Get NoData() As Boolean

    NoData = m_PrintInfo.NoData
    
End Property

Public Property Let NoData(Param As Boolean)

    m_PrintInfo.NoData = Param
    
End Property

