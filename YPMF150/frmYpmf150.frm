VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmYpmf150 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   3105
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf150.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   12150
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdLogin 
         Caption         =   "�J�ÔN�����ƒS���҂̕ύX"
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9900
         TabIndex        =   9
         Top             =   180
         Visible         =   0   'False
         Width           =   675
         _Version        =   262145
         _ExtentX        =   1191
         _ExtentY        =   635
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "2;20"
         Contents        =   "frmYpmf150.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�J�ÔN����"
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
         LabelWidth      =   75
         LabelHeight     =   25
         LabelLeft       =   11
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   8
         Left            =   3480
         TabIndex        =   13
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�S����"
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
         LabelWidth      =   45
         LabelHeight     =   25
         LabelLeft       =   26
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
      Begin VB.Label lblPname 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "�m�m�m�m�m�m�m�m�m�m"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   4980
         TabIndex        =   11
         Top             =   180
         Width           =   1905
      End
      Begin VB.Label lblOdate 
         Alignment       =   2  '��������
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "9999/12/31"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1620
         TabIndex        =   10
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   7
      Top             =   2280
      Width           =   12015
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "��ʸر(F8)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   2
         rPic.top        =   4
         rPic.right      =   0
         rPic.bottom     =   0
         rText.left      =   10
         rText.top       =   8
         rText.right     =   103
         rText.bottom    =   27
         Picture         =   "frmYpmf150.frx":0D13
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10200
         TabIndex        =   3
         Top             =   180
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
         Picture         =   "frmYpmf150.frx":0D2F
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   8460
         TabIndex        =   2
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "���(F12)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   9
         rPic.top        =   8
         rPic.right      =   16
         rPic.bottom     =   16
         rText.left      =   31
         rText.top       =   8
         rText.right     =   106
         rText.bottom    =   27
         Picture         =   "frmYpmf150.frx":0E89
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   1635
      Left            =   60
      TabIndex        =   6
      Top             =   660
      Width           =   12015
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   12240
      TabIndex        =   0
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf150.frx":0F9B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf150.frx":1009
      Key             =   "frmYpmf150.frx":1027
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
      Left            =   12240
      TabIndex        =   4
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf150.frx":106B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf150.frx":10D9
      Key             =   "frmYpmf150.frx":10F7
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
   Begin imText6Ctl.imText imtIcode_Kana_Focus1 
      Height          =   75
      Left            =   15225
      TabIndex        =   5
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf150.frx":113B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf150.frx":11A9
      Key             =   "frmYpmf150.frx":11C7
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
Attribute VB_Name = "frmYpmf150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ځ@�I�@�@�F
'���@���@�@�F��ʃN���A�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    '���̓`�F�b�N
    If DoValidationChecks() = False Then Exit Sub
    '����p���[�N�쐬
    If MakeWork() = False Then Exit Sub
    '����v���r���[
    If ActiveReportPrint(0) = False Then Exit Sub

    Exit Sub
    
cmdExecute_Click_Err:

    Call MsgBox("���s�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�I���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdLogin_Click()

    On Error GoTo cmdLogin_Click_Err

    g_blnLoginOK = False
    g_strPcode = Trim(cboPcode.Text)
    g_strPname = lblPname.Caption
    g_strOdate = Trim(lblOdate.Caption)
    frmLogin.Show vbModal
    If g_blnLoginOK = True Then
        lblOdate.Caption = g_strOdate
        cboPcode.Text = g_strPcode
        lblPname.Caption = g_strPname
        Call FieldsClear(0)
    End If
    Unload frmLogin
    
    Exit Sub

cmdLogin_Click_Err:

    Call MsgBox("�J�ÔN�����ƒS���҂̕ύX�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdLogin_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
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
            cmdClear.SetFocus
            DoEvents
            Call cmdClear_Click
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "�����W�v�\"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    Call FieldsClear(0)
    
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set g_clsAdoSQL = Nothing
    Set g_clsReg = Nothing
    End
    
    Exit Sub
    
Form_Unload_Err:

    Call MsgBox("�t�H�[���A�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")
    End

End Sub

'�ځ@�I�@�@�F��ʃN���A
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F0�F�S���
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("��ʃN���A�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

'�ځ@�I�@�@�F���̓`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "�J�ÔN��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "���̓`�F�b�N")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("���̓`�F�b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function
Private Function MakeWork() As Boolean
    
    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    
    Dim curSkeep As Currency                '�o�i�҈ێ��Ǘ���
    Dim curSrate As Currency                '�o�i�Ҏ萔����
    Dim intSfraction As Integer             '�o�i�Ғ[������
    Dim curSRounding As Currency            '�o�i�ҊۂߒP��
    Dim curTaxRate As Currency              '����ŗ�
    
    Dim curUriage As Currency
    Dim curPrice As Currency
    Dim curCharge As Currency
    Dim curTax As Currency
    Dim curKeep As Currency
    Dim curGTotal As Currency
    
    On Error GoTo MakeWork_Err
    
    MakeWork = False
    
    Screen.MousePointer = vbHourglass
    
    '������
    curSkeep = 0        '�o�i�҈ێ��Ǘ���
    intSfraction = 0    '�o�i�Ғ[������
    curSrate = 0        '�o�i�Ҏ萔����
    curSRounding = 0    '�o�i�ҊۂߒP��
    curTaxRate = 0      '����ŗ�
    
    '�ݒ�}�X�^�I�[�v��
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        If Not IsNull(adoMT010.Fields("Skeep")) Then curSkeep = adoMT010.Fields("Skeep")
        If Not IsNull(adoMT010.Fields("Sfraction")) Then intSfraction = adoMT010.Fields("Sfraction")
        If Not IsNull(adoMT010.Fields("Srate")) Then curSrate = adoMT010.Fields("Srate")
        If Not IsNull(adoMT010.Fields("SRounding")) Then curSRounding = adoMT010.Fields("SRounding")
    End If
    adoMT010.Close
    
    '����ŗ��擾
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, g_strOdate)

    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF033"
    g_clsAdoAccess.Connection.Execute strSQL
    
    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF033"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '�����f�[�^�I�[�v��
    strSQL = "SELECT * FROM DT030" & _
             " WHERE Odate = '" & lblOdate.Caption & "'"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        wkRecordset.AddNew
        wkRecordset.Fields("Odate") = adoRecordset1.Fields("Odate")
        wkRecordset.Fields("Onum") = adoRecordset1.Fields("Onum")
        wkRecordset.Fields("Sname") = adoRecordset1.Fields("Sname")

        '������
        curUriage = 0
        curPrice = 0
        curCharge = 0
        curTax = 0
        curKeep = 0
        curGTotal = 0
                
        '�������׃f�[�^�I�[�v��
        strSQL = "SELECT * FROM DT031" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Onum = " & adoRecordset1.Fields("Onum")
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset2.EOF
            '������z�����ʁ~����P��
            curUriage = curUriage + (CCur(adoRecordset2.Fields("Qty")) * CCur(adoRecordset2.Fields("Price1")))
            '�d�����z�����ʁ~�d���P��
            curPrice = curPrice + (CCur(adoRecordset2.Fields("Qty")) * CCur(adoRecordset2.Fields("Price2")))
            
            adoRecordset2.MoveNext
        Loop
        adoRecordset2.Close
        
        wkRecordset.Fields("Uriage") = curUriage
        wkRecordset.Fields("Price") = curPrice

        '�萔��
        If IsNull(adoRecordset1.Fields("ChargeDiv")) = False And adoRecordset1.Fields("ChargeDiv") = 1 Then
            curCharge = curPrice * curSrate / 100
            '�؂�̂�
            If IsNull(adoRecordset1.Fields("FixDiv")) = False And adoRecordset1.Fields("FixDiv") = 1 Then
                curCharge = Global_Rounding(curCharge, intSfraction, curSRounding)
            Else
                curCharge = Global_Rounding(curCharge, intSfraction, 1)
            End If
        End If
        '����Ōv�Z
        If IsNull(adoRecordset1.Fields("TaxDiv")) = False And adoRecordset1.Fields("TaxDiv") = 1 Then
            '�؂�̂�
            If IsNull(adoRecordset1.Fields("FixDiv")) = False And adoRecordset1.Fields("FixDiv") = 1 Then
                curTax = Global_Get_Tax(curPrice - curCharge, curTaxRate, intSfraction, curSRounding)
            Else
                curTax = Global_Get_Tax(curPrice - curCharge, curTaxRate, intSfraction, 1)
            End If
        End If
        '�ێ��Ǘ���
        If IsNull(adoRecordset1.Fields("KeepDiv")) = False And adoRecordset1.Fields("KeepDiv") = 1 Then
            curKeep = curSkeep
        End If
        '�����v�����v�|�萔���{����Ł|�ێ��Ǘ���
        curGTotal = curPrice - curCharge + curTax - curKeep

        wkRecordset.Fields("Charge") = curCharge
        wkRecordset.Fields("Total") = curPrice - curCharge
        wkRecordset.Fields("Tax") = curTax
        wkRecordset.Fields("Keep") = curKeep
        wkRecordset.Fields("GTotal") = curGTotal
        wkRecordset.Update
            
        adoRecordset1.MoveNext
    Loop
        
    adoRecordset1.Close
    wkRecordset.Requery     '�o�O�h�~
    wkRecordset.Close
    
    Screen.MousePointer = vbDefault
    
    MakeWork = True
    
    Exit Function
    
MakeWork_Err:

    MakeWork = False
    Screen.MousePointer = vbDefault
    Call MsgBox("���[�N�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork_Err")
    
End Function

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cmdExecute.SetFocus

End Sub

'�ځ@�I�@�@�FActiveReport�̈��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F0:�v���r���[ 1:���
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf033
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "�����W�v�\"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_Ypmf033"
        .Caption = "�����W�v�\"
        If .PrintActiveReport(0) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End With

    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    ActiveReportPrint = True
    
    Exit Function
    
ActiveReportPrint_Err:

    ActiveReportPrint = False
    Screen.MousePointer = vbDefault
    Call MsgBox("���s�N���b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReportPrint_Err")
    
End Function

'�ځ@�I�@�@�F�t�B�[���h�̃Z�b�g
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�O�R
'�X�V�����@�F
'
Private Function FieldsSet() As Boolean
    
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    Screen.MousePointer = vbDefault
    
    FieldsSet = True
    
    Exit Function

FieldsSet_Err:

    FieldsSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�t�B�[���h�Z�b�g�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsSet_Err")

End Function

