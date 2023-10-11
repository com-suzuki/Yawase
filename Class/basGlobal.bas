Attribute VB_Name = "basGlobal"
Option Explicit

'�X���[�v�֐�
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetVersion Lib "kernel32" () As Long

'API��NumLock��On�ɂ���
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'API�ŃL�[���͂��s��
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const VK_NUMLOCK = &H90
Public Const VK_CAPSLOCK = &H14
Public Const KEYEVENTF_KEYUP = &H2
Public Const WM_KEYDOWN = &H100     '�L�[�_�E��
Public Const VK_TAB = &H9           'TAB
Public Const VK_RETURN = &HD        'Enter
Public Const VK_SHIFT = &H10        'Shift

'�v���O�����I���܂őҋ@
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As Long, ByVal lpCommandLine As Long, ByVal IDProcess As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal lpdExitCode As Long, hHandle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ACTIVE = &H103

'SHBrowseForFolderq�Ŏg�p����\����
Private Type BROWSEINFO
    hwndOwner As Long       '�eWindow�������
    pidlRoot As Long        'ٰ�̫���
    pszDisplayName As Long
    lpszTitle As String     '�޲�۸ނɕ\������ү����
    ulFlags As Long         '��߼��
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'ٰ�̫��ޒ萔
Private Const CSIDL_DESKTOP = &H0           '�޽�į��
Private Const CSIDL_PROGRAMS = &H2          '��۸���
Private Const CSIDL_CONTROLS = &H3          '���۰�����
Private Const CSIDL_PRINTERS = &H4          '������
Private Const CSIDL_PERSONAL = &H5          '�߰���
Private Const CSIDL_FAVORITES = &H6         '�ޯ�ϰ�
Private Const CSIDL_STARTUP = &H7           '���ı���
Private Const CSIDL_RECENT = &H8            '[�ŋߎg����̧��]
Private Const CSIDL_SENDTO = &H9            '[����]
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB         '[����]�ƭ�
Private Const CSIDL_DESKTOPDIRECTORY = &H10 '�޽�į��
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOO = &H13           'Network Neighborhood
Private Const CSIDL_FONTS = &H14            '̫��
Private Const CSIDL_TEMPLATES = &H15        'Shell New

'����̫���(ϲ���߭���A���۰����ٓ�)��I�������Ȃ�
Private Const BIF_BROWSEFORCOMPUTER = 1
'[̫��ނ̎Q��]�޲�۸ނ��Ăяo��API
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFO As BROWSEINFO) As Long
'SHBrowseForFolder�œ���ꂽ�l����̫��ނ��߽���擾����API
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'SHBrowseForFolder�œ���ꂽ�l����؂��J������API
Private Declare Function SHFree Lib "shell32" Alias "#195" (ByVal pidl As Long) As Long
'�e���|�����t�@�C���̂��߂Ɏw�肳��Ă���p�X���擾����֐��̐錾
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'�N���X���ƃE�B���h�E�����w�肳�ꂽ������ƈ�v����E�B���h�E�̃n���h�����擾����֐��̐錾
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'�E�B���h�E���A�C�R��������Ă��邩�ǂ������f����֐��̐錾
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
'�E�B���h�E�𕜌����ăA�N�e�B�u������֐��̐錾
Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
'�w�肵���E�B���h�E����Ԏ�O�Ɏ����Ă���֐��̐錾
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

'INI�t�@�C������֐��錾
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type RECT
    left As Long    'Window��X���W
    top As Long     'Window��Y���W
    right As Long   'Window�̉E�[�̍��W
    bottom As Long  'Window�̒�ɂ����镔���̍��W
End Type
Private Const HWND_TOP = 0           '��O�ɾ��
Private Const HWND_BOTTOM = 1        '���ɾ��
Private Const HWND_TOPMOST = -1      '��Ɏ�O�ɾ��
Private Const HWND_NOTOPMOST = -2    '��Ɏ�O�A����
Private Const SWP_SHOWWINDOW = &H40  '�\������

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'�ځ@�I�@�@�FWindows�̃o�[�W��������
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Function IsWindowsNT() As Boolean

    IsWindowsNT = IIf(GetVersion() And &H80000000, False, True)

End Function

'�ځ@�I�@�@�FSendKeys�iSendKey("{TAB}")���Q��A�����Ď��s�����NumLock���͂����o�O����j
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Sub Global_SendKeys(ByVal OFORM As Object, ByVal CHARLNG As Long)

    Dim varState As Variant

    'NumLock���擾
    varState = GetKeyState(VK_NUMLOCK)

    'Next Field��
    Call PostMessage(OFORM.hwnd, WM_KEYDOWN, CHARLNG, 0)
    
    'NumLock On
    If varState <> 0 And GetKeyState(VK_NUMLOCK) = 0 Then
        Call keybd_event(VK_NUMLOCK, 0, 0, 0)
        Call keybd_event(VK_NUMLOCK, 0, KEYEVENTF_KEYUP, 0)
    End If

End Sub

'�ځ@�I�@�@�F�����_�ȉ��l�̌ܓ�
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_Round(ByVal dblArg1 As Double) As Double

    If dblArg1 >= 0 Then
        dblArg1 = CCur(dblArg1) + 0.5
    Else
        dblArg1 = CCur(dblArg1) - 0.5
    End If
    Global_Round = Fix(CCur(dblArg1))

End Function

'�ځ@�I�@�@�F�����_�ȉ��؂�グ
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_RoundUp(ByVal dblArg1 As Double) As Double

    If dblArg1 >= 0 Then
        dblArg1 = CCur(dblArg1) + 0.9999
    Else
        dblArg1 = CCur(dblArg1) - 0.9999
    End If
    Global_RoundUp = Fix(CCur(dblArg1))

End Function

'�ځ@�I�@�@�F�O���A�v���P�[�V�������I������܂őҋ@����
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_Exec_Application(ByRef strCmdline As String, ByRef intWindowstyle As Integer) As Boolean
    
    Dim hShell As Long
    Dim hProc As Long
    Dim lExit As Long
    Dim bret As Long

    On Error Resume Next

    hShell = Shell(strCmdline, intWindowstyle)
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hShell)
    Do
        GetExitCodeProcess hProc, lExit
        DoEvents
    Loop While lExit = STILL_ACTIVE
    
    bret = CloseHandle(hProc)

    Global_Exec_Application = True

    Exit Function
 
End Function

'�ځ@�I�@�@�F���̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FlngMonth�F�ΏۂƂȂ�N���@dblDiff�F��������i�O�j
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_Get_PassingMonth(ByVal lngMonth As Long, ByVal intDiff As Integer) As Long

    Dim intYear As Integer
    Dim intMonth As Integer
    Dim varDate As Variant
    
    On Error Resume Next
    
    '�N�ƌ��ɕ���
    intYear = left(lngMonth, 4)
    intMonth = right(lngMonth, 2)
    '���t�����߂�
    varDate = DateSerial(intYear, intMonth + intDiff, 1)
    
    Global_Get_PassingMonth = CLng(Format(varDate, "yyyymm"))
  
End Function

'�ځ@�I�@�@�F���̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FlngMonth�F�ΏۂƂȂ�N�����@dblDiff�F������i�O�j
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_Get_PassingDay(ByVal lngYyyymmdd As Long, ByVal intDiff As Integer) As Long

    Dim intYear As Integer
    Dim intMonth As Integer
    Dim intDay As Integer
    Dim varDate As Variant
    
    On Error Resume Next
    
    '�N�����ɕ���
    intYear = left(lngYyyymmdd, 4)
    intMonth = Mid(lngYyyymmdd, 5, 2)
    intDay = right(lngYyyymmdd, 2)
    
    '���t�����߂�
    varDate = DateSerial(intYear, intMonth, intDay + intDiff)
    
    Global_Get_PassingDay = CLng(Format(varDate, "yyyymmdd"))
  
End Function

'�ځ@�I�@�@�F���̍ŏI���擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FlngMonth�F�ΏۂƂȂ�N��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_Get_MonthLastDay(ByVal lngMonth As Long) As Integer

    Dim intYear As Integer
    Dim intMonth As Integer
    Dim varDate As Variant
    
    On Error Resume Next
    
    '�N�ƌ��ɕ���
    intYear = left(lngMonth, 4)
    intMonth = right(lngMonth, 2)
    '���t�����߂�
    varDate = DateSerial(intYear, intMonth + 1, 1 - 1)
    
    Global_Get_MonthLastDay = Day(varDate)
    
End Function

'�ځ@�I�@�@�FUnicode��Ansi
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FstrArg�F������@intByte�F�o�C�g��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_LeftB_Ansi(ByRef strArg As String, ByRef intByte As Integer) As String

    On Error Resume Next

    If strArg = "" Or intByte = 0 Then
        Global_LeftB_Ansi = ""
    End If

    Global_LeftB_Ansi = StrConv(LeftB(StrConv(strArg, vbFromUnicode), intByte), vbUnicode)

End Function

'�ځ@�I�@�@�FUnicode��Ansi
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FstrArg�F������@intByte�F�o�C�g��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_RightB_Ansi(ByRef strArg As String, ByRef intByte As Integer) As String

    On Error Resume Next

    If strArg = "" Or intByte = 0 Then
        Global_RightB_Ansi = ""
    End If

    Global_RightB_Ansi = StrConv(RightB(StrConv(strArg, vbFromUnicode), intByte), vbUnicode)

End Function

'�ځ@�I�@�@�FUnicode��Ansi
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FstrArg�F������@intByte�F�o�C�g��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_MidB_Ansi(ByRef strArg As String, ByRef intByte1 As Integer, ByRef intByte2 As Integer) As String

    On Error Resume Next

    If strArg = "" Or intByte1 = 0 Or intByte2 = 0 Then
        Global_MidB_Ansi = ""
    End If

    Global_MidB_Ansi = StrConv(MidB(StrConv(strArg, vbFromUnicode), intByte1, intByte2), vbUnicode)

End Function

'�ځ@�I�@�@�FUnicode��Ansi
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FstrArg�F������@intByte�F�o�C�g��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_LenB_Ansi(ByRef strArg As String) As Integer

    On Error Resume Next

    If strArg = "" Then
        Global_LenB_Ansi = 0
    End If

    Global_LenB_Ansi = LenB(StrConv(strArg, vbFromUnicode))

End Function

'�ځ@�I�@�@�F�����̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FlngMonth1�F�ΏۂƂȂ�N���P lngMonth2�F�ΏۂƂȂ�N���Q
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function Global_Get_DiffMonth(ByVal lngMonth1 As Long, ByVal lngMonth2 As Long) As Integer

    Dim strDate1 As String
    Dim strDate2 As String
    
    On Error Resume Next
    
    strDate1 = left(Format(lngMonth1, "000000"), 4) & "/" & right(Format(lngMonth1, "000000"), 2) & "/01"
    strDate2 = left(Format(lngMonth2, "000000"), 4) & "/" & right(Format(lngMonth2, "000000"), 2) & "/01"
    
    Global_Get_DiffMonth = Abs(DateDiff("m", Format(CDate(strDate1), "yyyy/mm/dd"), Format(CDate(strDate2), "yyyy/mm/dd")))
  
End Function

'�ځ@�I�@�@�F�t�H���_�I���_�C�A���O�\��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function OpenSelectFolderDialog(ByRef hwnd As Long) As String
    
    Dim typBrowseInfo As BROWSEINFO
    Dim lngFoldPointer As Long
    Dim strPathName As String

    On Error GoTo OpenSelectFolderDialog_Err

    OpenSelectFolderDialog = ""

    With typBrowseInfo
        '�e�E�C���h�E��ݒ�
        .hwndOwner = hwnd
        '���[�g�t�H���_��ݒ�
        .pidlRoot = 0
        .lpszTitle = "�t�H���_�I��"
        '����t�H���_��I�������Ȃ�
        .ulFlags = BIF_BROWSEFORCOMPUTER
    End With

    '[�t�H���_�̎Q��]�_�C�A���O���Ăяo��
    lngFoldPointer = SHBrowseForFolder(typBrowseInfo)
    If lngFoldPointer = 0 Then Exit Function

    '�\��Null�������Z�b�g
    strPathName = String$(256, vbNullChar)
    'SHBrowseForFolder�œ���ꂽ�l����t�H���_�̃p�X���擾
    Call SHGetPathFromIDList(lngFoldPointer, strPathName)

    '���蓖�Ă�ꂽ���������J��
    Call SHFree(lngFoldPointer)

    If Trim(strPathName) <> "" Then
        OpenSelectFolderDialog = strPathName
    End If
    
    Exit Function
    
OpenSelectFolderDialog_Err:

    OpenSelectFolderDialog = ""
    Call MsgBox("�t�H���_�I���_�C�A���O�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "OpenSelectFolderDialog_Err")

End Function

'�ځ@�I�@�@�Fyyyy/mm/dd(������)��yyyymmdd(���l)�ɕϊ�����
'���@���@�@�F�s�����t�`�F�b�N�Ȃ�
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�T
'�X�V�����@�F
'
Public Function Global_Get_NumericDay(ByRef strYyyymmdd As String) As Long

    Dim varBuff As Variant
    Dim strDay As String
    Dim intIndex1 As Integer
    
    On Error Resume Next
    
    varBuff = Split(strYyyymmdd, "/")
    strDay = ""
    For intIndex1 = 0 To 2
        strDay = strDay & varBuff(intIndex1)
    Next intIndex1
    
    Global_Get_NumericDay = CLng(strDay)
  
End Function

'�ځ@�I�@�@�Fyyyymmdd(���l)��yyyy/mm/dd(������)�ɕϊ�����
'���@���@�@�F�s�����t�`�F�b�N�Ȃ�
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�T
'�X�V�����@�F
'
Public Function Global_Get_StringDay(ByRef lngYyyymmdd As Long) As String

    Dim strDay As String
    
    On Error Resume Next
    
    strDay = left$(Format(lngYyyymmdd, "00000000"), 4) & "/"
    strDay = strDay & Mid$(Format(lngYyyymmdd, "00000000"), 5, 2) & "/"
    strDay = strDay & right$(Format(lngYyyymmdd, "00000000"), 2)
    
    Global_Get_StringDay = strDay
  
End Function

'�ځ@�I�@�@�Fyyyy mm dd��yyyy/mm/dd�ɂ���
'���@���@�@�F�s�����t�`�F�b�N�Ȃ�
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�T
'�X�V�����@�F
'
Public Function Global_StrToDate(ByRef strYyyy As String, ByRef strMm As String, ByRef strDd As String) As String

    On Error Resume Next
    Global_StrToDate = Trim(strYyyy) & "/" & Format(strMm, "00") & "/" & Format(strDd, "00")
  
End Function

'�ځ@�I�@�@�F�d���N���̃`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F�N���X���^�^�C�g���^���s�`���t�@�C���p�X�^���s�t�@�C��
'�߂�l�@�@�F���큁True�^�G���[��False
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�P�^�O�R�^�O�U
'�X�V�����@�F
'
Public Function Global_IsLoad(strClassName As String, strAppTitle As String, strBinPath As String, strExe As String) As Boolean

    Dim lHandleWindow As Long
    Dim varResponse  As Variant

    Global_IsLoad = True

    On Error GoTo Global_IsLoad_Err

    lHandleWindow = FindWindow(strClassName, strAppTitle)
    If lHandleWindow <> 0 Then
        If IsIconic(lHandleWindow) <> False Then
            varResponse = OpenIcon(lHandleWindow)
        Else
            varResponse = BringWindowToTop(lHandleWindow)
        End If
    Else
        If strBinPath <> "" And strExe <> "" Then
            varResponse = Shell(strBinPath & strExe, vbNormalFocus)
        End If
    End If

    Exit Function

Global_IsLoad_Err:
    
    Global_IsLoad = False

End Function

'�ځ@�I�@�@�FINI�t�@�C������̃f�[�^�擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�P�^�O�R�^�O�T
'�X�V�����@�F
'
Public Function Global_GetIni(ByVal vstrAppname As String, ByVal vstrKeyword As String, ByVal vstrIniFile As String) As String
    
    Dim strResult   As String * 1024
    Dim intTemp     As Integer
    Dim intLen      As Integer
    Dim strFileName As String
   
    On Error Resume Next
   
    intTemp = GetPrivateProfileString(vstrAppname, vstrKeyword, "", strResult, Len(strResult), vstrIniFile)
    Global_GetIni = left$(strResult, intTemp)

End Function

'�ځ@�I�@�@�FINI�t�@�C���փf�[�^�X�V
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�P�^�O�R�^�O�T
'�X�V�����@�F
'
Public Function Global_SetIni(ByVal vstrAppname As String, ByVal vstrKeyword As String, ByVal vstrKeyVal As String, ByVal vstrIniFile As String) As Integer
    
    Dim intTemp As Integer

    On Error Resume Next

    intTemp = WritePrivateProfileString(vstrAppname, vstrKeyword, vstrKeyVal, vstrIniFile)
    Global_SetIni = intTemp

End Function

'�ځ@�I�@�@�F���t�̃`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�T
'�X�V�����@�F
'
Public Function Global_IsDate(ByRef strYyyy As String, ByRef strMm As String, ByRef strDd As String) As Boolean

    On Error Resume Next
    Global_IsDate = IsDate(Global_StrToDate(strYyyy, strMm, strDd))
  
End Function

'�ځ@�I�@�@�F�t�H�[������Ɏ�O�ɕ\����ݒ�^��������
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�T
'�X�V�����@�F
'
Public Function Global_SetFormTop(lngHWnd As Long, bolIsTop As Boolean)
    
    Dim mLeft As Long
    Dim mTop As Long
    Dim mWidth As Long
    Dim mHeight As Long
    Dim lpRect As RECT

    On Error GoTo Global_SetFormTop_Err

    Global_SetFormTop = GetWindowRect(lngHWnd, lpRect)

    mLeft = lpRect.left
    mTop = lpRect.top
    mWidth = lpRect.right - lpRect.left
    mHeight = lpRect.bottom - lpRect.top
    
    If bolIsTop Then
        Global_SetFormTop = SetWindowPos(lngHWnd, HWND_TOPMOST, mLeft, mTop, mWidth, mHeight, SWP_SHOWWINDOW)
    Else
        Global_SetFormTop = SetWindowPos(lngHWnd, HWND_NOTOPMOST, mLeft, mTop, mWidth, mHeight, SWP_SHOWWINDOW)
    End If

    Exit Function

Global_SetFormTop_Err:

    Call MsgBox("�t�H�[������Ɏ�O�ɕ\����ݒ�^�����G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_SetFormTop_Err")

End Function

Public Function Global_StrNull(ByVal varString As Variant, Optional ByVal strDefault As String = "") As String

    On Error Resume Next

    Global_StrNull = IIf(IsNull(varString) = True, strDefault, varString)

End Function

Public Function Global_GetPcName() As String

    Dim strBuff As String
    Dim intPos As Integer
    
    Const MAX_COMPUTERNAME_LENGTH = 255

    On Error Resume Next

    strBuff = Space(MAX_COMPUTERNAME_LENGTH + 1)
    GetComputerName strBuff, MAX_COMPUTERNAME_LENGTH

    'Chr$(0)�ȍ~�̕������폜
    intPos = InStr(strBuff, Chr$(0))
    If intPos > 0 Then
        strBuff = left$(strBuff, intPos - 1)
    End If
    
    Global_GetPcName = strBuff
    
End Function
