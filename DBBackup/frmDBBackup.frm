VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmDBBackup 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   2655
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDBBackup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   8325
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame fra 
      Height          =   2535
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   8175
      Begin VB.ComboBox cboBackupDrive 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1860
         TabIndex        =   6
         Text            =   "cboBackupDrive"
         Top             =   780
         Width           =   735
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   6360
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "�I��(F9)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SizePicture     =   -1  'True
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   14
         rPic.top        =   6
         rPic.right      =   21
         rPic.bottom     =   21
         rText.left      =   43
         rText.top       =   8
         rText.right     =   109
         rText.bottom    =   27
         Picture         =   "frmDBBackup.frx":0CFA
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   4560
         TabIndex        =   1
         Top             =   1920
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "���s(F12)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SizePicture     =   -1  'True
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   9
         rPic.top        =   6
         rPic.right      =   21
         rPic.bottom     =   21
         rText.left      =   34
         rText.top       =   8
         rText.right     =   109
         rText.bottom    =   27
         Picture         =   "frmDBBackup.frx":0E54
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   5
         Left            =   180
         TabIndex        =   7
         Top             =   780
         Width           =   1575
         _Version        =   262145
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�o�b�N�A�b�v��"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         TopMargin       =   0
         LabelTop        =   0
         LabelWidth      =   89
         LabelHeight     =   25
         LabelLeft       =   8
         BottomMargin    =   0
         RightMargin     =   0
         Spacing         =   0
         AutoAdjust      =   -1  'True
         BorderEffect    =   1
         BorderStyle     =   1
         LabelAutoSize   =   1
         LabelPosition   =   0
         ToolTip         =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "�����̑S�Ă̏������I�����Ă�����s���Ă��������B"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   7755
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   8640
      TabIndex        =   0
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmDBBackup.frx":12A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDBBackup.frx":1314
      Key             =   "frmDBBackup.frx":1332
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   0
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   10
      LengthAsByte    =   -1
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin imText6Ctl.imText imtFocusEnd 
      Height          =   135
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmDBBackup.frx":1376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDBBackup.frx":13E4
      Key             =   "frmDBBackup.frx":1402
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   0
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   10
      LengthAsByte    =   -1
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
End
Attribute VB_Name = "frmDBBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsReg As New clsReg

Private Declare Function GetLogicalDrives Lib "KERNEL32" () As Long

Private Sub cboBackupDrive_DropDown()

    Dim Drives As Long
    Dim I As Long
    Dim Bit As Long
    Dim strBuff As String

    On Error GoTo cboBackupDrive_DropDown_Err

    Bit = 1
    
    strBuff = cboBackupDrive.Text
    cboBackupDrive.Clear

    '�h���C�u�ꗗ�擾
    Drives = GetLogicalDrives()

    '���b�Z�[�W�쐬
    For I = Asc("A") To Asc("Z")
        If (Drives And Bit) <> 0 Then
            cboBackupDrive.AddItem Chr(I) & ":"
        End If
        Bit = Bit * 2
    Next

    cboBackupDrive.Text = strBuff

    Exit Sub

cboBackupDrive_DropDown_Err:

    Call MsgBox("�h���b�v�_�E�����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBackupDrive_DropDown_Err")

End Sub

Private Sub cboBackupDrive_GotFocus()

    cboBackupDrive.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub cboBackupDrive_LostFocus()

    cboBackupDrive.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�Q�Q
'�X�V�����@�F
'
Private Sub cmdExecute_Click()
    
    Dim ServerObj  As New SQLDMO.SQLServer
    Dim BackObj    As New SQLDMO.Backup
    
    On Error GoTo cmdExecute_Click_Err

    If MsgBox("���s���܂����H", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    ServerObj.Connect m_clsReg.Server, m_clsReg.UID, m_clsReg.PWD
    
    '�o�b�N�A�b�v�̑Ώ�DB
    BackObj.Database = m_clsReg.DBName
    '�o�b�N�A�b�v�t�@�C���̕ۑ��ꏊ
    BackObj.Files = Trim(cboBackupDrive.Text) & "\" & Format(Now(), "yyyymmdd") & ".bak"
    '�f�[�^�x�[�X�S�̂��o�b�N�A�b�v����
    BackObj.Action = SQLDMOBackup_Database
    ''�o�b�N�A�b�v��������s����
    BackObj.SQLBackup ServerObj
    
    Screen.MousePointer = vbDefault
    
    Call MsgBox("�o�b�N�A�b�v�������I�����܂����B", vbOKOnly + vbInformation, "���")

    Exit Sub

cmdExecute_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("���s�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�I���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�Q�Q
'�X�V�����@�F
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�Q�Q
'�X�V�����@�F
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Form_KeyDown_Err
    
    '���^�[���L�[�Ŏ��̃R���g���[���փt�H�[�J�X�ړ�
    If KeyCode = vbKeyReturn And Shift <> vbShiftMask Then
        KeyCode = 0
        Global_SendKeys Me, VK_TAB
        Exit Sub
    End If
    
    '�V���[�g�J�b�g�L�[�̊��蓖��
    Select Case KeyCode
        Case vbKeyF1
        Case vbKeyF2
        Case vbKeyF3
        Case vbKeyF4
        Case vbKeyF5
        Case vbKeyF6
        Case vbKeyF7
        Case vbKeyF8
        Case vbKeyF9
            cmdExit.SetFocus
            DoEvents
            Call cmdExit_Click
        Case vbKeyF10
        Case vbKeyF11
        Case vbKeyF12
            cmdExecute.SetFocus
            DoEvents
            Call cmdExecute_Click
        Case vbKeyF2
        Case vbKeyHome
        Case vbKeyPageUp
        Case vbKeyPageDown
    End Select

    Exit Sub

Form_KeyDown_Err:

    Call MsgBox("�t�H�[���L�[�_�E�����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[�����[�h��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�Q�Q
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "�f�[�^�x�[�X�̃o�b�N�A�b�v"

    '�d���N���̃`�F�b�N
    If App.PrevInstance = True Then
        Unload Me
        End
    End If
        
    '���W�X�g���ǂݍ���
    m_clsReg.RegKey = REG_KEY
    If m_clsReg.ReadReg = False Then
        Unload Me
        End
    End If
    
    cboBackupDrive.Text = m_clsReg.BackUpDrive
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("�t�H�[�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    End

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[���A�����[�h��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�Q�Q
'�X�V�����@�F
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set m_clsReg = Nothing
    End
    
    Exit Sub
    
Form_Unload_Err:

    Call MsgBox("�t�H�[���A�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")
    End

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cmdExecute.SetFocus

End Sub


