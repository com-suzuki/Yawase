VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmYpmf260 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   3270
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf260.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   9930
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   540
      TabIndex        =   19
      Top             =   3600
      Width           =   6975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   9735
      Begin VB.Frame Frame2 
         BorderStyle     =   0  '�Ȃ�
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   780
         Width           =   2175
         Begin VB.OptionButton optDiv 
            Caption         =   "�Г�"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   3
            Top             =   120
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optDiv 
            Caption         =   "�ЊO"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1020
            TabIndex        =   4
            Top             =   120
            Width           =   915
         End
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�W�v�N��"
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
         LabelWidth      =   60
         LabelHeight     =   25
         LabelLeft       =   18
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
      Begin imText6Ctl.imText txtYear 
         Height          =   420
         Left            =   1680
         TabIndex        =   1
         Top             =   300
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   741
         Caption         =   "frmYpmf260.frx":0CFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf260.frx":0D68
         Key             =   "frmYpmf260.frx":0D86
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   4
         LengthAsByte    =   0
         Text            =   "9999"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   3
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText txtMonth 
         Height          =   420
         Left            =   3120
         TabIndex        =   2
         Top             =   300
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf260.frx":0DBA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf260.frx":0E28
         Key             =   "frmYpmf260.frx":0E46
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   3
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   12
         Left            =   180
         TabIndex        =   16
         Top             =   900
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "��@��"
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
         LabelWidth      =   40
         LabelHeight     =   25
         LabelLeft       =   28
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
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "���@�L"
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
         LabelWidth      =   40
         LabelHeight     =   25
         LabelLeft       =   28
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
      Begin imText6Ctl.imText txtTokki 
         Height          =   885
         Left            =   1680
         TabIndex        =   5
         Top             =   1500
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   1561
         Caption         =   "frmYpmf260.frx":0E7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf260.frx":0EE8
         Key             =   "frmYpmf260.frx":0F06
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
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
         MultiLine       =   -1
         ScrollBars      =   2
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   1
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   1
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3780
         TabIndex        =   14
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Caption         =   "�N"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   10
      Top             =   2460
      Width           =   9735
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   7860
         TabIndex        =   7
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
         Picture         =   "frmYpmf260.frx":0F4A
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   6120
         TabIndex        =   6
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
         rPic.left       =   10
         rPic.top        =   8
         rPic.right      =   16
         rPic.bottom     =   16
         rText.left      =   30
         rText.top       =   8
         rText.right     =   105
         rText.bottom    =   27
         Picture         =   "frmYpmf260.frx":10A4
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdCsv 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "CSV�o��(F8)"
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
         rPic.left       =   10
         rPic.top        =   8
         rPic.right      =   0
         rPic.bottom     =   0
         rText.left      =   8
         rText.top       =   8
         rText.right     =   105
         rText.bottom    =   27
         Picture         =   "frmYpmf260.frx":11B6
      End
   End
   Begin imText6Ctl.imText imtFocusEnd 
      Height          =   135
      Left            =   10320
      TabIndex        =   8
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf260.frx":11D2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf260.frx":1240
      Key             =   "frmYpmf260.frx":125E
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
      TabIndex        =   9
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf260.frx":12A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf260.frx":1310
      Key             =   "frmYpmf260.frx":132E
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
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   10140
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf260.frx":1372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf260.frx":13E0
      Key             =   "frmYpmf260.frx":13FE
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
Attribute VB_Name = "frmYpmf260"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsAdoSQL As New clsAdoCore
Private m_clsAdoAccess As New clsAdoCore
Private m_clsReg As New clsReg

Private Type typCSV_Record
    Field001 As String
    Field002 As Double
    Field003 As Double
    Field004 As Double
    Field005 As Double
    Field006 As Double
    Field007 As Double
    Field008 As Double
    Field009 As Double
End Type

Private Sub cmdCsv_Click()

    Dim strSQL As String
    Dim wkRecordset As New ADODB.Recordset
    Dim intFreefile1 As Integer
    Dim Csv_Rec As typCSV_Record

    On Error GoTo cmdCsv_Click_Err
    
    txtFileName.Text = ""
    
    With CommonDialog1
        .DialogTitle = "csv̧�ق��w��"
        .FileName = ""
        .CancelError = False
        .Filter = "csv̧�� (*.csv)|*.csv|���ׂĂ�̧�� (*.*)|*.*|"
        '.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer
        '�R�����_�C�A���O�{�b�N�X���J��
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        '�t�@�C�����擾
        txtFileName.Text = .FileName
    End With
    
    If Trim(txtFileName.Text) = "" Then Exit Sub
    '���Ƀt�@�C��������ꍇ
    If Dir(txtFileName.Text) <> "" Then
        If MsgBox("�㏑�����܂����H", vbInformation + vbYesNo, "") = vbNo Then Exit Sub
    End If
    
    '���̓`�F�b�N
    If DoValidationChecks() = False Then Exit Sub
    '����p���[�N�쐬
    If MakePrintWork() = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    '�N�G���[�I�[�v��
    strSQL = "SELECT * FROM QWK_YPMF260"
    wkRecordset.Open strSQL, m_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    If wkRecordset.EOF = True Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '�t�@�C���쐬
    intFreefile1 = FreeFile
    Open txtFileName.Text For Output As intFreefile1
    
    '�^�C�g��
    If optDiv(0).Value = True Then
        Write #intFreefile1, "���", "�s��", "�s�O", "����", "�k��", "�֐�", "�֓�", "���̑�", "������z"
    Else
        Write #intFreefile1, "���", "�s��", "�s�O", "����", "�k��", "�֐�", "�֓�", "���̑�"
    End If
    
    Do While Not wkRecordset.EOF
        Csv_Rec.Field001 = IIf(IsNull(wkRecordset.Fields("Kind_Name")), "", Trim(wkRecordset.Fields("Kind_Name")))
        Csv_Rec.Field002 = IIf(IsNull(wkRecordset.Fields("In_City")), 0, wkRecordset.Fields("In_City"))
        Csv_Rec.Field003 = IIf(IsNull(wkRecordset.Fields("Out_City")), 0, wkRecordset.Fields("Out_City"))
        Csv_Rec.Field004 = IIf(IsNull(wkRecordset.Fields("In_Pref")), 0, wkRecordset.Fields("In_Pref"))
        Csv_Rec.Field005 = IIf(IsNull(wkRecordset.Fields("Hokuriku")), 0, wkRecordset.Fields("Hokuriku"))
        Csv_Rec.Field006 = IIf(IsNull(wkRecordset.Fields("Kansai")), 0, wkRecordset.Fields("Kansai"))
        Csv_Rec.Field007 = IIf(IsNull(wkRecordset.Fields("Kantou")), 0, wkRecordset.Fields("Kantou"))
        Csv_Rec.Field008 = IIf(IsNull(wkRecordset.Fields("Etc")), 0, wkRecordset.Fields("Etc"))
        Csv_Rec.Field009 = IIf(IsNull(wkRecordset.Fields("Uriage_Kingaku")), 0, wkRecordset.Fields("Uriage_Kingaku"))
    
        '��������
        If optDiv(0).Value = True Then
            Write #intFreefile1, Csv_Rec.Field001, Csv_Rec.Field002, Csv_Rec.Field003, Csv_Rec.Field004, Csv_Rec.Field005, _
                                 Csv_Rec.Field006, Csv_Rec.Field007, Csv_Rec.Field008, Csv_Rec.Field009
        Else
            Write #intFreefile1, Csv_Rec.Field001, Csv_Rec.Field002, Csv_Rec.Field003, Csv_Rec.Field004, Csv_Rec.Field005, _
                                 Csv_Rec.Field006, Csv_Rec.Field007, Csv_Rec.Field008
        End If
        
        wkRecordset.MoveNext
    Loop
    
    Close intFreefile1
    DoEvents
    
    Screen.MousePointer = vbDefault

    Call MsgBox("�I�����܂����B", vbOKOnly + vbInformation, "")
    DoEvents

    Exit Sub

cmdCsv_Click_Err:

    Close
    Screen.MousePointer = vbDefault
    Call MsgBox("CSV�o�̓N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdCsv_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    '���̓`�F�b�N
    If DoValidationChecks() = False Then Exit Sub
    '����p���[�N�쐬
    If MakePrintWork() = False Then Exit Sub
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
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
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
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
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
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "�A�ؗ��ʓ��������\�P"

    '�d���N���̃`�F�b�N
    If App.PrevInstance = True Then
        End
    End If

    '�t�H�[���N���A
    txtYear.Text = Format(Now(), "yyyy")
    txtMonth.Text = Format(Now(), "m")
    
    '���W�X�g���ǂݍ���
    m_clsReg.RegKey = REG_KEY
    If m_clsReg.ReadReg = False Then
        End
    End If

    '�f�[�^�x�[�X�ڑ�
    With m_clsAdoSQL
        .Provider = adoSQLServer
        .Server = m_clsReg.Server
        .DBName = m_clsReg.DBName
        .UID = m_clsReg.UID
        .PWD = m_clsReg.PWD
        .CommandTimeOut = m_clsReg.CommandTimeOut
        If .Connect = False Then
            End
        End If
    End With
    With m_clsAdoAccess
        .Provider = adoAccess
        .DBName = m_clsReg.LDatabase & "\" & m_clsReg.LDBName
        If .Connect = False Then
            End
        End If
    End With
    
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
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set m_clsAdoSQL = Nothing
    Set m_clsAdoAccess = Nothing
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

    txtYear.SetFocus

End Sub

Private Sub txtMonth_GotFocus()

    txtMonth.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtMonth_LostFocus()

    txtMonth.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtTokki_GotFocus()

    txtTokki.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtTokki_LostFocus()

    txtTokki.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtYear_GotFocus()

    txtYear.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtYear_LostFocus()

    txtYear.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F����p���[�N�쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim adoMT060 As New ADODB.Recordset
    Dim adoDT011 As New ADODB.Recordset
    Dim adoDT021 As New ADODB.Recordset
    Dim adoDT031 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim adoCmd As New ADODB.Command
    Dim adoParam As ADODB.Parameter
    Dim strYyyymm1 As String
    Dim strYyyymm2 As String

    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    '�W�v�N��(YYYY/MM)
    strYyyymm1 = txtYear.Text & "/" & Format(txtMonth.Text, "00")
    '�W�v�N��(YYYYMM)
    strYyyymm2 = txtYear.Text & Format(txtMonth.Text, "00")
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF260"
    m_clsAdoAccess.Connection.Execute strSQL

    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF260"
    wkRecordset.Open strSQL, m_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '���i�敪�}�X�^�I�[�v��
    strSQL = "{call sp_MT060;1}"
    adoMT060.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT060.EOF = True Then
        Call MsgBox("���i�敪�}�X�^���ݒ肳��Ă��܂���B", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If
    
    With frmCount
        .fpProgressBar1.Max = adoMT060.RecordCount
        .fpProgressBar1.Value = 1
        .Show
        Me.Enabled = False
    End With
    
    Do While Not adoMT060.EOF
        '���[�N�쐬
        wkRecordset.AddNew
        wkRecordset.Fields("Yyyymm") = txtYear.Text & "�N" & Format(txtMonth.Text, "00") & "��"
        wkRecordset.Fields("Kind_Div1") = left$(Format(adoMT060.Fields("Dcode"), "00"), 1)
        wkRecordset.Fields("Kind_Div2") = adoMT060.Fields("Dcode")
        wkRecordset.Fields("Kind_Name") = adoMT060.Fields("Dname")
            
'********** ���ח� **********
        
        wkRecordset.Fields("In_City") = 0
        wkRecordset.Fields("Out_City") = 0
        
        '********** ��t�f�[�^ **********
        
        '��t���׃f�[�^
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2601;1('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2601;3('" & strYyyymm1 & "')}"
        End If
        adoDT011.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT011.EOF
            If Not IsNull(adoDT011.Fields("Div")) And Not IsNull(adoDT011.Fields("Qty")) Then
                If adoDT011.Fields("Div") = TIKU_DIV_ON Then
                    wkRecordset.Fields("In_City") = CCur(wkRecordset.Fields("In_City")) + CCur(adoDT011.Fields("Qty"))
                ElseIf adoDT011.Fields("Div") = TIKU_DIV_OFF Then
                    wkRecordset.Fields("Out_City") = CCur(wkRecordset.Fields("Out_City")) + CCur(adoDT011.Fields("Qty"))
                Else
                    wkRecordset.Fields("In_City") = CCur(wkRecordset.Fields("In_City")) + CCur(adoDT011.Fields("Qty"))
                End If
            End If
            adoDT011.MoveNext
        Loop
        adoDT011.Close
        '��t���׃f�[�^(�ݐ�)
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2601;2('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2601;4('" & strYyyymm1 & "')}"
        End If
        adoDT011.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT011.EOF
            If Not IsNull(adoDT011.Fields("Div")) And Not IsNull(adoDT011.Fields("Qty")) Then
                If adoDT011.Fields("Div") = TIKU_DIV_ON Then
                    wkRecordset.Fields("In_City") = CCur(wkRecordset.Fields("In_City")) + CCur(adoDT011.Fields("Qty"))
                ElseIf adoDT011.Fields("Div") = TIKU_DIV_OFF Then
                    wkRecordset.Fields("Out_City") = CCur(wkRecordset.Fields("Out_City")) + CCur(adoDT011.Fields("Qty"))
                Else
                    wkRecordset.Fields("In_City") = CCur(wkRecordset.Fields("In_City")) + CCur(adoDT011.Fields("Qty"))
                End If
            End If
            adoDT011.MoveNext
        Loop
        adoDT011.Close
    
        '********** �����f�[�^ **********
    
        '�������׃f�[�^
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2608;1('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2608;3('" & strYyyymm1 & "')}"
        End If
        adoDT031.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT031.EOF
            If Not IsNull(adoDT031.Fields("Div")) And Not IsNull(adoDT031.Fields("Qty")) Then
                If adoDT031.Fields("Div") = TIKU_DIV_ON Then
                    wkRecordset.Fields("In_City") = CCur(wkRecordset.Fields("In_City")) + CCur(adoDT031.Fields("Qty"))
                ElseIf adoDT031.Fields("Div") = TIKU_DIV_OFF Then
                    wkRecordset.Fields("Out_City") = CCur(wkRecordset.Fields("Out_City")) + CCur(adoDT031.Fields("Qty"))
                Else
                    wkRecordset.Fields("In_City") = CCur(wkRecordset.Fields("In_City")) + CCur(adoDT031.Fields("Qty"))
                End If
            End If
            adoDT031.MoveNext
        Loop
        adoDT031.Close
        '�������׃f�[�^(�ݐ�)
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2608;2('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2608;4('" & strYyyymm1 & "')}"
        End If
        adoDT031.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT031.EOF
            If Not IsNull(adoDT031.Fields("Div")) And Not IsNull(adoDT031.Fields("Qty")) Then
                If adoDT031.Fields("Div") = TIKU_DIV_ON Then
                    wkRecordset.Fields("In_City") = CCur(wkRecordset.Fields("In_City")) + CCur(adoDT031.Fields("Qty"))
                ElseIf adoDT031.Fields("Div") = TIKU_DIV_OFF Then
                    wkRecordset.Fields("Out_City") = CCur(wkRecordset.Fields("Out_City")) + CCur(adoDT031.Fields("Qty"))
                Else
                    wkRecordset.Fields("In_City") = CCur(wkRecordset.Fields("In_City")) + CCur(adoDT031.Fields("Qty"))
                End If
            End If
            adoDT031.MoveNext
        Loop
        adoDT031.Close
    
'********** �o�ח� **********
    
        wkRecordset.Fields("In_Pref") = 0
        wkRecordset.Fields("Hokuriku") = 0
        wkRecordset.Fields("Kansai") = 0
        wkRecordset.Fields("Kantou") = 0
        wkRecordset.Fields("Etc") = 0
    
        '********** �����f�[�^ **********
    
        '�������׃f�[�^
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2602;1('" & strYyyymm2 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2602;3('" & strYyyymm2 & "')}"
        End If
        adoDT021.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT021.EOF
            If Not IsNull(adoDT021.Fields("Qty")) Then
                '�n����擾
                Select Case Global_Postal_to_Pref(m_clsReg.LDatabase, adoDT021.Fields("Post"))
                    Case 0:
                        wkRecordset.Fields("In_Pref") = CCur(wkRecordset.Fields("In_Pref")) + CCur(adoDT021.Fields("Qty"))
                    Case 1:
                        wkRecordset.Fields("Hokuriku") = CCur(wkRecordset.Fields("Hokuriku")) + CCur(adoDT021.Fields("Qty"))
                    Case 2:
                        wkRecordset.Fields("Kansai") = CCur(wkRecordset.Fields("Kansai")) + CCur(adoDT021.Fields("Qty"))
                    Case 3:
                        wkRecordset.Fields("Kantou") = CCur(wkRecordset.Fields("Kantou")) + CCur(adoDT021.Fields("Qty"))
                    Case 4:
                        wkRecordset.Fields("Etc") = CCur(wkRecordset.Fields("Etc")) + CCur(adoDT021.Fields("Qty"))
                End Select
            End If
            adoDT021.MoveNext
        Loop
        adoDT021.Close
        '�������׃f�[�^(�ݐ�)
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2602;2('" & strYyyymm2 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2602;4('" & strYyyymm2 & "')}"
        End If
        adoDT021.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT021.EOF
            If Not IsNull(adoDT021.Fields("Qty")) Then
                '�n����擾
                Select Case Global_Postal_to_Pref(m_clsReg.LDatabase, adoDT021.Fields("Post"))
                    Case 0:
                        wkRecordset.Fields("In_Pref") = CCur(wkRecordset.Fields("In_Pref")) + CCur(adoDT021.Fields("Qty"))
                    Case 1:
                        wkRecordset.Fields("Hokuriku") = CCur(wkRecordset.Fields("Hokuriku")) + CCur(adoDT021.Fields("Qty"))
                    Case 2:
                        wkRecordset.Fields("Kansai") = CCur(wkRecordset.Fields("Kansai")) + CCur(adoDT021.Fields("Qty"))
                    Case 3:
                        wkRecordset.Fields("Kantou") = CCur(wkRecordset.Fields("Kantou")) + CCur(adoDT021.Fields("Qty"))
                    Case 4:
                        wkRecordset.Fields("Etc") = CCur(wkRecordset.Fields("Etc")) + CCur(adoDT021.Fields("Qty"))
                End Select
            End If
            adoDT021.MoveNext
        Loop
        adoDT021.Close

        '********** �����f�[�^ **********

        '��t���׃f�[�^�i�������j
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2604;1('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2604;3('" & strYyyymm1 & "')}"
        End If
        adoDT011.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT011.EOF
            If Not IsNull(adoDT011.Fields("Qty")) Then
                '�n����擾
                Select Case Global_Postal_to_Pref(m_clsReg.LDatabase, adoDT011.Fields("Post"))
                    Case 0:
                        wkRecordset.Fields("In_Pref") = CCur(wkRecordset.Fields("In_Pref")) + CCur(adoDT011.Fields("Qty"))
                    Case 1:
                        wkRecordset.Fields("Hokuriku") = CCur(wkRecordset.Fields("Hokuriku")) + CCur(adoDT011.Fields("Qty"))
                    Case 2:
                        wkRecordset.Fields("Kansai") = CCur(wkRecordset.Fields("Kansai")) + CCur(adoDT011.Fields("Qty"))
                    Case 3:
                        wkRecordset.Fields("Kantou") = CCur(wkRecordset.Fields("Kantou")) + CCur(adoDT011.Fields("Qty"))
                    Case 4:
                        wkRecordset.Fields("Etc") = CCur(wkRecordset.Fields("Etc")) + CCur(adoDT011.Fields("Qty"))
                End Select
            End If
            adoDT011.MoveNext
        Loop
        adoDT011.Close
        '��t���׃f�[�^(�ݐ�)�i�������j
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2604;2('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2604;4('" & strYyyymm1 & "')}"
        End If
        adoDT011.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT011.EOF
            If Not IsNull(adoDT011.Fields("Qty")) Then
                '�n����擾
                Select Case Global_Postal_to_Pref(m_clsReg.LDatabase, adoDT011.Fields("Post"))
                    Case 0:
                        wkRecordset.Fields("In_Pref") = CCur(wkRecordset.Fields("In_Pref")) + CCur(adoDT011.Fields("Qty"))
                    Case 1:
                        wkRecordset.Fields("Hokuriku") = CCur(wkRecordset.Fields("Hokuriku")) + CCur(adoDT011.Fields("Qty"))
                    Case 2:
                        wkRecordset.Fields("Kansai") = CCur(wkRecordset.Fields("Kansai")) + CCur(adoDT011.Fields("Qty"))
                    Case 3:
                        wkRecordset.Fields("Kantou") = CCur(wkRecordset.Fields("Kantou")) + CCur(adoDT011.Fields("Qty"))
                    Case 4:
                        wkRecordset.Fields("Etc") = CCur(wkRecordset.Fields("Etc")) + CCur(adoDT011.Fields("Qty"))
                End Select
            End If
            adoDT011.MoveNext
        Loop
        adoDT011.Close
    
        '�������׃f�[�^
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2606;1('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2606;3('" & strYyyymm1 & "')}"
        End If
        adoDT031.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT031.EOF
            If Not IsNull(adoDT031.Fields("Qty")) Then
                '�n����擾
                Select Case Global_Postal_to_Pref(m_clsReg.LDatabase, adoDT031.Fields("Post"))
                    Case 0:
                        wkRecordset.Fields("In_Pref") = CCur(wkRecordset.Fields("In_Pref")) + CCur(adoDT031.Fields("Qty"))
                    Case 1:
                        wkRecordset.Fields("Hokuriku") = CCur(wkRecordset.Fields("Hokuriku")) + CCur(adoDT031.Fields("Qty"))
                    Case 2:
                        wkRecordset.Fields("Kansai") = CCur(wkRecordset.Fields("Kansai")) + CCur(adoDT031.Fields("Qty"))
                    Case 3:
                        wkRecordset.Fields("Kantou") = CCur(wkRecordset.Fields("Kantou")) + CCur(adoDT031.Fields("Qty"))
                    Case 4:
                        wkRecordset.Fields("Etc") = CCur(wkRecordset.Fields("Etc")) + CCur(adoDT031.Fields("Qty"))
                End Select
            End If
            adoDT031.MoveNext
        Loop
        adoDT031.Close
        '�������׃f�[�^(�ݐ�)
        If adoMT060.Fields("Dcode") <> 99 Then
            strSQL = "{call sp_YPMF2606;2('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
        Else
            strSQL = "{call sp_YPMF2606;4('" & strYyyymm1 & "')}"
        End If
        adoDT031.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT031.EOF
            If Not IsNull(adoDT031.Fields("Qty")) Then
                '�n����擾
                Select Case Global_Postal_to_Pref(m_clsReg.LDatabase, adoDT031.Fields("Post"))
                    Case 0:
                        wkRecordset.Fields("In_Pref") = CCur(wkRecordset.Fields("In_Pref")) + CCur(adoDT031.Fields("Qty"))
                    Case 1:
                        wkRecordset.Fields("Hokuriku") = CCur(wkRecordset.Fields("Hokuriku")) + CCur(adoDT031.Fields("Qty"))
                    Case 2:
                        wkRecordset.Fields("Kansai") = CCur(wkRecordset.Fields("Kansai")) + CCur(adoDT031.Fields("Qty"))
                    Case 3:
                        wkRecordset.Fields("Kantou") = CCur(wkRecordset.Fields("Kantou")) + CCur(adoDT031.Fields("Qty"))
                    Case 4:
                        wkRecordset.Fields("Etc") = CCur(wkRecordset.Fields("Etc")) + CCur(adoDT031.Fields("Qty"))
                End Select
            End If
            adoDT031.MoveNext
        Loop
        adoDT031.Close
    
'********** ������z **********
    
        wkRecordset.Fields("Uriage_Kingaku") = 0
        
        If optDiv(0).Value = True Then
            '�������׃f�[�^
            If adoMT060.Fields("Dcode") <> 99 Then
                strSQL = "{call sp_YPMF2603;1('" & strYyyymm2 & "'," & adoMT060.Fields("Dcode") & ")}"
            Else
                strSQL = "{call sp_YPMF2603;3('" & strYyyymm2 & "')}"
            End If
            adoDT021.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT021.EOF = False Then
                If Not IsNull(adoDT021.Fields("Price")) Then
                    wkRecordset.Fields("Uriage_Kingaku") = CCur(wkRecordset.Fields("Uriage_Kingaku")) + CCur(adoDT021.Fields("Price"))
                End If
            End If
            adoDT021.Close
            '�������׃f�[�^(�ݐ�)
            If adoMT060.Fields("Dcode") <> 99 Then
                strSQL = "{call sp_YPMF2603;2('" & strYyyymm2 & "'," & adoMT060.Fields("Dcode") & ")}"
            Else
                strSQL = "{call sp_YPMF2603;4('" & strYyyymm2 & "')}"
            End If
            adoDT021.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT021.EOF = False Then
                If Not IsNull(adoDT021.Fields("Price")) Then
                    wkRecordset.Fields("Uriage_Kingaku") = CCur(wkRecordset.Fields("Uriage_Kingaku")) + CCur(adoDT021.Fields("Price"))
                End If
            End If
            adoDT021.Close
            
            '��t���׃f�[�^(������)
            If adoMT060.Fields("Dcode") <> 99 Then
                strSQL = "{call sp_YPMF2605;1('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
            Else
                strSQL = "{call sp_YPMF2605;3('" & strYyyymm1 & "')}"
            End If
            adoDT011.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT011.EOF = False Then
                If Not IsNull(adoDT011.Fields("Price")) Then
                    wkRecordset.Fields("Uriage_Kingaku") = CCur(wkRecordset.Fields("Uriage_Kingaku")) + CCur(adoDT011.Fields("Price"))
                End If
            End If
            adoDT011.Close
            '��t���׃f�[�^(�ݐ�)(������)
            If adoMT060.Fields("Dcode") <> 99 Then
                strSQL = "{call sp_YPMF2605;2('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
            Else
                strSQL = "{call sp_YPMF2605;4('" & strYyyymm1 & "')}"
            End If
            adoDT011.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT011.EOF = False Then
                If Not IsNull(adoDT011.Fields("Price")) Then
                    wkRecordset.Fields("Uriage_Kingaku") = CCur(wkRecordset.Fields("Uriage_Kingaku")) + CCur(adoDT011.Fields("Price"))
                End If
            End If
            adoDT011.Close
            
            '�������׃f�[�^
            If adoMT060.Fields("Dcode") <> 99 Then
                strSQL = "{call sp_YPMF2607;1('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
            Else
                strSQL = "{call sp_YPMF2607;3('" & strYyyymm1 & "')}"
            End If
            adoDT031.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT031.EOF = False Then
                If Not IsNull(adoDT031.Fields("Price")) Then
                    wkRecordset.Fields("Uriage_Kingaku") = CCur(wkRecordset.Fields("Uriage_Kingaku")) + CCur(adoDT031.Fields("Price"))
                End If
            End If
            adoDT031.Close
            '�������׃f�[�^(�ݐ�)
            If adoMT060.Fields("Dcode") <> 99 Then
                strSQL = "{call sp_YPMF2607;2('" & strYyyymm1 & "'," & adoMT060.Fields("Dcode") & ")}"
            Else
                strSQL = "{call sp_YPMF2607;4('" & strYyyymm1 & "')}"
            End If
            adoDT031.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT031.EOF = False Then
                If Not IsNull(adoDT031.Fields("Price")) Then
                    wkRecordset.Fields("Uriage_Kingaku") = CCur(wkRecordset.Fields("Uriage_Kingaku")) + CCur(adoDT031.Fields("Price"))
                End If
            End If
            adoDT031.Close
        End If
    
        wkRecordset.Fields("Subtotal_Name") = Global_CirculationTitle(wkRecordset.Fields("Kind_Div1"))
        If Trim(wkRecordset.Fields("Subtotal_Name")) <> "" Then
            wkRecordset.Fields("Subtotal_Name") = Trim(wkRecordset.Fields("Subtotal_Name")) & " �v"
        End If
        
        wkRecordset.Update
    
        adoMT060.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    
    wkRecordset.Requery
    wkRecordset.Close
    
    
    MakePrintWork = True
    
MakePrintWork_Exit:
    
    Me.Enabled = True
    Unload frmCount
    
    Screen.MousePointer = vbDefault
    
    Exit Function

MakePrintWork_Cancel:

    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    MakePrintWork = False
    Call MsgBox("������[�N�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function

'�ځ@�I�@�@�FActiveReport�̈��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F0:�v���r���[ 1:���
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf260
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "�A�ؗ��ʓ��������\�P"
        .objReport = objRpt
        .Connection = m_clsAdoAccess.Connection
        .Caption = "�A�ؗ��ʓ��������\�P"
        If .PrintActiveReport(intFlg) = False Then
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
    Call MsgBox("ActiveReport�̈���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReportPrint_Err")
    
End Function

'�ځ@�I�@�@�F���̓`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtYear.Text) = "" Then
        txtYear.SetFocus
        strErrMsg = "�W�v�N������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtMonth.Text) = "" Then
        txtMonth.SetFocus
        strErrMsg = "�W�v�N������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If CInt(txtYear.Text) < 1900 Or CInt(txtYear.Text) > 2099 Then
        txtYear.SetFocus
        strErrMsg = "����N�S���œ��͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If CInt(txtMonth.Text) < 1 Or CInt(txtMonth.Text) > 12 Then
        txtMonth.SetFocus
        strErrMsg = "�������W�v�N������͂��Ă��������B"
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


