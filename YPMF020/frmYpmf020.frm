VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmYpmf020 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   10260
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf020.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   15030
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdYpmf020Clear 
      Caption         =   "���̓��[�N�̃N���A"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   180
      Width           =   2715
   End
   Begin VB.Frame fraHeader 
      Height          =   615
      Left            =   60
      TabIndex        =   64
      Top             =   1260
      Width           =   14895
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   60
         TabIndex        =   65
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "��ԍ�"
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
      Begin imText6Ctl.imText txtHnum 
         Height          =   345
         Left            =   1560
         TabIndex        =   2
         Top             =   180
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   609
         Caption         =   "frmYpmf020.frx":0CFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":0D68
         Key             =   "frmYpmf020.frx":0D86
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
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   10
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWW"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
   End
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   4140
      TabIndex        =   56
      Top             =   600
      Width           =   9675
      Begin VB.CommandButton cmdLogin 
         Caption         =   "�J�ÔN�����ƒS���҂̕ύX"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   6960
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9900
         TabIndex        =   57
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
         Contents        =   "frmYpmf020.frx":0DBA
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   60
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
         TabIndex        =   61
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
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraRecordSelector 
      Height          =   615
      Left            =   7920
      TabIndex        =   46
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmdDataMove 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1680
         Picture         =   "frmYpmf020.frx":0DD3
         Style           =   1  '���̨���
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1140
         Picture         =   "frmYpmf020.frx":0F1D
         Style           =   1  '���̨���
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   600
         Picture         =   "frmYpmf020.frx":1067
         Style           =   1  '���̨���
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   60
         Picture         =   "frmYpmf020.frx":11B1
         Style           =   1  '���̨���
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   45
      Top             =   9480
      Width           =   14895
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   30
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
         Picture         =   "frmYpmf020.frx":12FB
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   13080
         TabIndex        =   32
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
         Picture         =   "frmYpmf020.frx":1317
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   11340
         TabIndex        =   31
         Top             =   180
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
         Picture         =   "frmYpmf020.frx":1471
      End
   End
   Begin VB.Frame fraSyori 
      Height          =   615
      Left            =   60
      TabIndex        =   39
      Top             =   0
      Width           =   7815
      Begin VB.OptionButton optSyori 
         Caption         =   "�O���o��"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6540
         Style           =   1  '���̨���
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "�m�F�\���"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5160
         Style           =   1  '���̨���
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   180
         Width           =   1395
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "��@��"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3960
         Style           =   1  '���̨���
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "�ρ@�X"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2760
         Style           =   1  '���̨���
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "�V�@�K"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1560
         Style           =   1  '���̨���
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   43
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�����敪"
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
   End
   Begin VB.Frame fraMeisai 
      Height          =   7635
      Left            =   60
      TabIndex        =   37
      Top             =   1860
      Width           =   14895
      Begin VB.CommandButton cmdUekisearch 
         Caption         =   "�A�ؖ��Ō���"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12060
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   540
         Width           =   2655
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "���ב}��(&I)"
         Height          =   375
         Left            =   7800
         TabIndex        =   21
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdQty 
         Caption         =   "�ύX"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1740
         Width           =   915
      End
      Begin VB.Frame fraDetail 
         Height          =   3135
         Left            =   9480
         TabIndex        =   71
         Top             =   120
         Visible         =   0   'False
         Width           =   5355
         Begin VB.CheckBox chkBdiv 
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2220
            TabIndex        =   24
            Top             =   1020
            Width           =   615
         End
         Begin VB.CheckBox chkSdiv 
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2220
            TabIndex        =   23
            Top             =   600
            Width           =   615
         End
         Begin VB.CheckBox chkWdiv 
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2220
            TabIndex        =   22
            Top             =   180
            Width           =   615
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   12
            Left            =   60
            TabIndex        =   72
            Top             =   180
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "�m�F�\�o�͍ς�"
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
            LabelWidth      =   105
            LabelHeight     =   25
            LabelLeft       =   16
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
            Index           =   13
            Left            =   60
            TabIndex        =   73
            Top             =   600
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "�o�i�`�[�o�͍ς�"
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
            LabelWidth      =   120
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
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   14
            Left            =   60
            TabIndex        =   74
            Top             =   1020
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "����`�[�o�͍ς�"
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
            LabelWidth      =   120
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
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   15
            Left            =   60
            TabIndex        =   75
            Top             =   1440
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "���吸�Z��"
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
            LabelWidth      =   90
            LabelHeight     =   25
            LabelLeft       =   23
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
            Index           =   16
            Left            =   60
            TabIndex        =   76
            Top             =   1860
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "�o�i�Ґ��Z��"
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
            LabelWidth      =   105
            LabelHeight     =   25
            LabelLeft       =   16
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
            Index           =   17
            Left            =   60
            TabIndex        =   77
            Top             =   2700
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "���͎���"
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
            LabelLeft       =   38
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
            Index           =   18
            Left            =   60
            TabIndex        =   78
            Top             =   2280
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "�S���҃R�[�h"
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
            LabelWidth      =   82
            LabelHeight     =   25
            LabelLeft       =   27
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
         Begin imNumber6Ctl.imNumber imnBnum 
            Height          =   375
            Left            =   2160
            TabIndex        =   25
            Top             =   1440
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   661
            Calculator      =   "frmYpmf020.frx":18C3
            Caption         =   "frmYpmf020.frx":18E3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf020.frx":1951
            Keys            =   "frmYpmf020.frx":196F
            Spin            =   "frmYpmf020.frx":19B9
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99
            MinValue        =   -99
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnSnum 
            Height          =   375
            Left            =   2160
            TabIndex        =   26
            Top             =   1860
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   661
            Calculator      =   "frmYpmf020.frx":19E1
            Caption         =   "frmYpmf020.frx":1A01
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf020.frx":1A6F
            Keys            =   "frmYpmf020.frx":1A8D
            Spin            =   "frmYpmf020.frx":1AD7
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99
            MinValue        =   -99
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   19
            Left            =   2820
            TabIndex        =   84
            Top             =   180
            Width           =   1815
            _Version        =   262145
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "���Z�s�ԍ�"
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
            LabelLeft       =   23
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
         Begin imNumber6Ctl.imNumber imnSline 
            Height          =   375
            Left            =   4680
            TabIndex        =   27
            Top             =   180
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   661
            Calculator      =   "frmYpmf020.frx":1AFF
            Caption         =   "frmYpmf020.frx":1B1F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf020.frx":1B8D
            Keys            =   "frmYpmf020.frx":1BAB
            Spin            =   "frmYpmf020.frx":1BF5
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99
            MinValue        =   -99
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2011496453
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   21
            Left            =   2820
            TabIndex        =   87
            Top             =   600
            Width           =   1815
            _Version        =   262145
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "�ύX�O�s�ԍ�"
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
            LabelWidth      =   90
            LabelHeight     =   25
            LabelLeft       =   15
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
         Begin imNumber6Ctl.imNumber imnOrgNum 
            Height          =   375
            Left            =   4680
            TabIndex        =   28
            Top             =   600
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   661
            Calculator      =   "frmYpmf020.frx":1C1D
            Caption         =   "frmYpmf020.frx":1C3D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf020.frx":1CAB
            Keys            =   "frmYpmf020.frx":1CC9
            Spin            =   "frmYpmf020.frx":1D13
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99
            MinValue        =   -99
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2011496453
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDetailPcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "99"
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
            Left            =   2160
            TabIndex        =   81
            Top             =   2280
            Width           =   405
         End
         Begin VB.Label lblItime 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "yyyy/mm/dd hh:mm:ss"
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
            Left            =   2160
            TabIndex        =   80
            Top             =   2700
            Width           =   2385
         End
         Begin VB.Label lblDetailPname 
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
            Left            =   2640
            TabIndex        =   79
            Top             =   2280
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "..."
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton cmdSearchDT011 
         Caption         =   "��t�\���猟��(&R)"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2580
         TabIndex        =   9
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "���׃R�s�[(&C)"
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear_Dst 
         Caption         =   "���׃N���A(N)"
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Top             =   2880
         Width           =   1575
      End
      Begin MSComctlLib.ListView lsvMeisai 
         Height          =   3795
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3300
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   6694
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "�s"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "��t�ԍ�"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��t�s�ԍ�"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "���i�R�[�h"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "�A�ؖ���"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "���@��"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "�������z"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "���喼��"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "�m�F�\�o�͋敪"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "�o�i�`�[�敪"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "����`�[�敪"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "���吸�Z��"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "�o�i�Ґ��Z��"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "���͎���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "�S���҃R�[�h"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "�S���Җ�"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Text            =   "���Z�s�ԍ�"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Text            =   "�����s�����t���O"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   19
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "�ύX�O�s�ԍ�"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "���׍폜(&D)"
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "���דo�^(&A)"
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   2880
         Width           =   1575
      End
      Begin imNumber6Ctl.imNumber imnQty 
         Height          =   495
         Left            =   1560
         TabIndex        =   12
         Top             =   1740
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   873
         Calculator      =   "frmYpmf020.frx":1D3B
         Caption         =   "frmYpmf020.frx":1D5B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   20.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":1DC9
         Keys            =   "frmYpmf020.frx":1DE7
         Spin            =   "frmYpmf020.frx":1E31
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,##0"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
         MinValue        =   -999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2011496453
         Value           =   999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imText6Ctl.imText txtIname 
         Height          =   525
         Left            =   1560
         TabIndex        =   11
         Top             =   1140
         Width           =   13095
         _Version        =   65536
         _ExtentX        =   23098
         _ExtentY        =   926
         Caption         =   "frmYpmf020.frx":1E59
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":1EC7
         Key             =   "frmYpmf020.frx":1EE5
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   0
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   40
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   1
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
         Index           =   3
         Left            =   60
         TabIndex        =   44
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�s�ԍ�"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   55
         Top             =   1740
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "���@��"
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
      Begin imNumber6Ctl.imNumber imnNo 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   180
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmYpmf020.frx":1F29
         Caption         =   "frmYpmf020.frx":1F49
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":1FB7
         Keys            =   "frmYpmf020.frx":1FD5
         Spin            =   "frmYpmf020.frx":201F
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   -99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   2011496453
         Value           =   99
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   9
         Left            =   60
         TabIndex        =   63
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "��t�ԍ�"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   11
         Left            =   60
         TabIndex        =   66
         Top             =   1140
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�A�ؖ�"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   4
         Left            =   4440
         TabIndex        =   67
         Top             =   1740
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�������z"
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
      Begin imNumber6Ctl.imNumber imnPrice 
         Height          =   495
         Left            =   5940
         TabIndex        =   14
         Top             =   1740
         Width           =   2595
         _Version        =   65536
         _ExtentX        =   4577
         _ExtentY        =   873
         Calculator      =   "frmYpmf020.frx":2047
         Caption         =   "frmYpmf020.frx":2067
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   20.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":20D5
         Keys            =   "frmYpmf020.frx":20F3
         Spin            =   "frmYpmf020.frx":212D
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999
         MinValue        =   -999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2011496453
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imText6Ctl.imText imtPnum 
         Height          =   480
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   847
         Caption         =   "frmYpmf020.frx":2155
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":21C3
         Key             =   "frmYpmf020.frx":21E1
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
         LengthAsByte    =   -1
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   6
         Left            =   5400
         TabIndex        =   68
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "��t�s�ԍ�"
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
      Begin imText6Ctl.imText imtPnumLine 
         Height          =   480
         Left            =   6900
         TabIndex        =   8
         Top             =   600
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   847
         Caption         =   "frmYpmf020.frx":2215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":2283
         Key             =   "frmYpmf020.frx":22A1
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   0
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
         LengthAsByte    =   -1
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
      Begin CSComboLib.CSComboBox cboBcode 
         Height          =   480
         Left            =   1560
         TabIndex        =   16
         Top             =   2280
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
         _ExtentY        =   847
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18.01
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;20"
         Contents        =   "frmYpmf020.frx":22E5
         Extended        =   -1  'True
         ListBoxWidth    =   650
         MaxLength       =   4
         Text            =   "9999"
         ValueCol        =   0
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   10
         Left            =   60
         TabIndex        =   69
         Top             =   2280
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "����R�[�h"
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
         LabelWidth      =   67
         LabelHeight     =   25
         LabelLeft       =   15
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
      Begin CSComboLib.CSComboBox cboIcode 
         Height          =   360
         Left            =   8040
         TabIndex        =   10
         Top             =   660
         Visible         =   0   'False
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         Enabled         =   0   'False
         ColDelim        =   ";"
         ColWidths       =   "2;20"
         Contents        =   "frmYpmf020.frx":22FE
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   5
         Text            =   "99999"
      End
      Begin imNumber6Ctl.imNumber imnPrice_Total 
         Height          =   375
         Left            =   10200
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   7140
         Width           =   2475
         _Version        =   65536
         _ExtentX        =   4366
         _ExtentY        =   661
         Calculator      =   "frmYpmf020.frx":2317
         Caption         =   "frmYpmf020.frx":2337
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":23A5
         Keys            =   "frmYpmf020.frx":23C3
         Spin            =   "frmYpmf020.frx":240D
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999
         MinValue        =   -999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2011496453
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   20
         Left            =   9780
         TabIndex        =   85
         Top             =   1740
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�����s����"
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
      Begin VB.CheckBox chkIdiv 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   26.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11340
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label lblPriceTani 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�~"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   20.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   8580
         TabIndex        =   86
         Top             =   1770
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "���z���v"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9060
         TabIndex        =   83
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Label lblBname 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2760
         TabIndex        =   70
         Top             =   2280
         Width           =   8505
      End
   End
   Begin VB.Frame fraKey 
      Height          =   675
      Left            =   60
      TabIndex        =   36
      Top             =   600
      Width           =   4035
      Begin VB.CheckBox chkAutoCode 
         Caption         =   "���ގ����̔�"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         Style           =   1  '���̨���
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   3420
         Picture         =   "frmYpmf020.frx":2435
         Style           =   1  '���̨���
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   180
         Width           =   555
      End
      Begin imText6Ctl.imText txtOcode 
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   180
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   635
         Caption         =   "frmYpmf020.frx":273F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf020.frx":27AD
         Key             =   "frmYpmf020.frx":27CB
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
         MaxLength       =   12
         LengthAsByte    =   0
         Text            =   "99999999999"
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
         Index           =   5
         Left            =   60
         TabIndex        =   38
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�����ԍ�"
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
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   15120
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf020.frx":280F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf020.frx":287D
      Key             =   "frmYpmf020.frx":289B
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
      Left            =   15120
      TabIndex        =   33
      Top             =   180
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf020.frx":28DF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf020.frx":294D
      Key             =   "frmYpmf020.frx":296B
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
      Height          =   135
      Left            =   15120
      TabIndex        =   34
      Top             =   3000
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf020.frx":29AF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf020.frx":2A1D
      Key             =   "frmYpmf020.frx":2A3B
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
   Begin imText6Ctl.imText imtIcode_Kana_Focus2 
      Height          =   135
      Left            =   15120
      TabIndex        =   35
      Top             =   3120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf020.frx":2A7F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf020.frx":2AED
      Key             =   "frmYpmf020.frx":2B0B
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
   Begin imText6Ctl.imText imtPnum_Focus1 
      Height          =   135
      Left            =   15120
      TabIndex        =   6
      Top             =   1980
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf020.frx":2B4F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf020.frx":2BBD
      Key             =   "frmYpmf020.frx":2BDB
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
   Begin imText6Ctl.imText imtPnum_Focus2 
      Height          =   135
      Left            =   15120
      TabIndex        =   7
      Top             =   2160
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf020.frx":2C1F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf020.frx":2C8D
      Key             =   "frmYpmf020.frx":2CAB
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
Attribute VB_Name = "frmYpmf020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_strPcName As String                        'PC��

Public m_strLastOcode As String                     '�Ō�ɓo�^���������ԍ�
Public m_blnTotalFlg As Boolean                     '���Z�t���O
Private m_curPriceTani As Currency                  '�������z�P��

Const AUTO_CODE = 1                                 '�R�[�h�̎����̔�
Const MAX_ROW = 20                                  '���ׂ̍ő�s��

Private Sub cboBcode_Click()

    Call cboBcode_Validate(False)

End Sub

Private Sub cboBcode_DropDown()

    Call MakecboBcode(cboBcode)

End Sub

Private Sub cboBcode_GotFocus()
   
    cboBcode.BackColor = FOCUS_STOP_COLOR
    cboBcode.Tag = cboBcode.Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboBcode_LostFocus()
   
    cboBcode.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboBcode_Validate(Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cboBcode_Validate_Err
    
    If Trim(cboBcode.Text) = "" Then
        lblBname.Caption = ""
        Exit Sub
    End If
    If IsNumeric(cboBcode.Text) = False Then
        cboBcode.Text = ""
        lblBname.Caption = ""
        Exit Sub
    End If
    If cboBcode.Tag = cboBcode.Text Then Exit Sub
    
    lblBname.Caption = ""
    
    With adoRecordset1
        '���Ӑ�}�X�^
        strSQL = "{call sp_MT070;2(" & Trim(cboBcode.Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Fdiv")) Then
                If .Fields("Fdiv") = BUSINESS_DIV_BUYER Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    lblBname.Caption = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If lblBname.Caption = "" Then cboBcode.Text = ""
    
    Exit Sub

cboBcode_Validate_Err:

    Call MsgBox("�t�H�[�J�X�ړ��O�G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBcode_Validate_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���ގ����̔ԃN���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub chkAutoCode_Click()

    On Error Resume Next

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then
         txtOcode.Text = AutoCodeSet()
         If txtOcode.Enabled Then txtOcode.SetFocus
    End If

End Sub

Private Sub chkIdiv_Click()

    If chkIdiv.Value = 1 Then
        imnPrice.Value = 0
        cboBcode.Text = ""
        lblBname.Caption = ""
    End If

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F��ʃN���A�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    On Error Resume Next
    If optSyori(0).Value = True Then
        If chkAutoCode.Value = 1 Then txtOcode.Text = AutoCodeSet()
    End If
    txtOcode.SetFocus

End Sub

Private Sub cmdClear_Dst_Click()

    Call FieldsClear(3)
    Call ListViewGetMaxRow
    imtPnum.SetFocus

End Sub

Private Sub cmdCopy_Click()

    Call ListViewGetMaxRow

    '���׃N���A
    chkWdiv.Value = 0
    chkSdiv.Value = 0
    chkBdiv.Value = 0
    imnBnum.Value = 0
    imnSnum.Value = 0
    lblDetailPcode.Caption = ""
    lblDetailPname.Caption = ""
    lblItime.Caption = ""

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���R�[�h�ړ��N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub cmdDataMove_Click(Index As Integer)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cmdDataMove_Click_Err

    Screen.MousePointer = vbHourglass

    With adoRecordset1
        strSQL = "SELECT * FROM DT020" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            Select Case Index
                Case 0:
                    .MoveFirst
                Case 1:
                    If Trim(txtOcode.Text) = "" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    .Find "Ocode = '" & Trim(txtOcode.Text) & "'"
                    If Not .EOF Then
                        .MovePrevious
                        If .EOF Or .BOF Then .MoveFirst
                    Else
                        .MoveFirst
                    End If
                Case 2:
                    If Trim(txtOcode.Text) = "" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    .Find "Ocode = '" & Trim(txtOcode.Text) & "'"
                    If Not .EOF Then
                        .MoveNext
                        If .EOF Or .BOF Then .MoveLast
                    Else
                        .MoveLast
                    End If
                Case 3:
                    .MoveLast
                Case Else
                    Exit Sub
            End Select
            Call FieldsClear(0)
            If FieldsSet(True, adoRecordset1) = False Then Exit Sub
        End If
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
cmdDataMove_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("���R�[�h�ړ��N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdDataMove_Click_Err")
    
End Sub

Private Sub cmdDel_Click()
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim intRow As Integer
    
    On Error GoTo cmdDel_Click_Err

    '���̓`�F�b�N
    If chkBdiv.Value = 1 Then
        Call MsgBox("���吸�Z����Ă��܂��B" & vbCrLf & "���̓`�[�̂��ׂĂ̔���Əo�i�҂̓`�[���ďo�͂��Ă��������B", vbOKOnly + vbCritical, "")
    End If
    
    If ListViewDelItem() = False Then Exit Sub
    If Trim(imtPnum.Text) <> "" And Trim(imtPnumLine.Text) <> "" Then
        '2004/01/26�@�����Ńt���O���͂����Ă��܂��ƃL�����Z�����ꂽ���ɍ���
        '��t�f�[�^�̃t���O���͂���
'        If DataDelete_DT011(imtPnum.Text, imtPnumLine.Text) = False Then Exit Sub
        
        '���͓r���̃��[�N�f�[�^�폜
        strSQL = "DELETE FROM YPMF020" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & imtPnum.Text & _
                 " AND Line = " & Format(imtPnumLine.Text, "####") & _
                 " AND PcName = '" & m_strPcName & "'"
        g_clsAdoSQL.Connection.Execute strSQL
        
        '�ύX�������̂�(2004/01/31�ǉ�)
        If optSyori(1).Value = True Then
            '�폜���[�N�ɒǉ�
            intRow = UBound(g_usrMeisaiDel) + 1
            ReDim Preserve g_usrMeisaiDel(intRow)
            
            With g_usrMeisaiDel(intRow)
                .Pnum = imtPnum.Text
                .PnumLine = imtPnumLine.Text
            End With
        End If
        
    End If
    Call Calc_Total
    Call FieldsClear(3)
    
    imtPnum.SetFocus

    Exit Sub

cmdDel_Click_Err:
    
     Call MsgBox("���׍폜�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdDel_Click_Err")

End Sub

Private Sub cmdDetail_Click()

    fraDetail.Visible = Not fraDetail.Visible

End Sub

Private Sub cmdEdit_Click()

    Dim strBuff As String
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim itmX As ListItem
    
    On Error GoTo cmdEdit_Click_Err

    If DoValidationChecks_Dst() = False Then Exit Sub
    
    
    '���X�g�r���[�̃f�[�^�����i�s�ԍ�����v����f�[�^����������폜�j
    Set itmX = lsvMeisai.FindItem(imnNo.Value, , , 0)
    If Not (itmX Is Nothing) Then
        '���͓r���̃��[�N�f�[�^�폜
        strSQL = "DELETE FROM YPMF020" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & itmX.SubItems(1) & _
                 " AND Line = " & Format(itmX.SubItems(2), "####") & _
                 " AND PcName = '" & m_strPcName & "'"
        g_clsAdoSQL.Connection.Execute strSQL
    End If
    Set itmX = Nothing
    
    
    If ListViewSetItem(imnNo.Value, 0) = False Then Exit Sub
    Call Calc_Total
    
    '���͓r���̃��[�N�f�[�^�ɒǉ�
    strSQL = "INSERT INTO YPMF020 (Odate,Pnum,Line,PcName)"
    strSQL = strSQL & " VALUES('" & lblOdate.Caption & "',"
    strSQL = strSQL & imtPnum.Text & ","
    strSQL = strSQL & Format(imtPnumLine.Text, "####") & ","
    strSQL = strSQL & "'" & m_strPcName & "')"
    g_clsAdoSQL.Connection.Execute strSQL
    
    strBuff = imtPnum.Text
    Call FieldsClear(3)
    imtPnum.Text = strBuff
    
    imtPnum.SetFocus

    Exit Sub
    
cmdEdit_Click_Err:

    Call MsgBox("���דo�^�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdEdit_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub cmdExecute_Click()

    On Error Resume Next

    If MsgBox("���s���܂����H", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    If optSyori(0).Value = True Or optSyori(1).Value = True Then
        '���̓`�F�b�N
        If DoValidationChecks() = False Then Exit Sub
        If DataUpdate() = False Then Exit Sub
        m_strLastOcode = txtOcode.Text
        
        '2005/09/01 �ύX�����m�F�\�_�C�A���O��\������
        If optSyori(0).Value = True Or optSyori(1).Value = True Then
            If MsgBox("�m�F�\��������܂����H", vbYesNo + vbQuestion, "") = vbYes Then
                frmPrintDialog.m_blnAutoPrint = True
                frmPrintDialog.Show vbModal
            End If
        End If
        
    ElseIf optSyori(2).Value = True Then
        '���̓`�F�b�N
        If DoValidationChecks() = False Then Exit Sub
        If DataDelete() = False Then Exit Sub
    End If
    
    '�t�B�[���h�N���A
    Call FieldsClear(0)

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then
        txtOcode.Text = AutoCodeSet()
        txtHnum.SetFocus
    Else
        txtOcode.SetFocus
    End If

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�I���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdInsert_Click()

    Call SetImeMode(ActiveControl.hwnd, 2)
    If DoValidationChecks_Dst() = False Then Exit Sub
    If ListViewInsItem() = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    imtPnum.SetFocus

End Sub

Private Sub cmdLogin_Click()

    Dim strSQL As String

    g_blnLoginOK = False
    g_strPcode = Trim(cboPcode.Text)
    g_strPname = lblPname.Caption
    g_strOdate = Trim(lblOdate.Caption)
    frmLogin.Show vbModal
    If g_blnLoginOK = True Then
        lblOdate.Caption = g_strOdate
        cboPcode.Text = g_strPcode
        lblPname.Caption = g_strPname
        If optSyori(0).Value = True Then
            Call optSyori_Click(0)
        Else
            optSyori(0).Value = True
        End If
        
        '���͓r���̃��[�N�f�[�^�폜
        strSQL = "DELETE FROM YPMF020" & _
                 " WHERE PcName = '" & m_strPcName & "'"
        g_clsAdoSQL.Connection.Execute strSQL
    End If
    Unload frmLogin
    
End Sub

Private Sub cmdQty_Click()

    On Error Resume Next

    imnQty.Enabled = Not imnQty.Enabled
    If imnQty.Enabled Then imnQty.SetFocus

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�����N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub cmdSearch_Click()

    frmSearch.Show vbModal

End Sub

Private Sub cmdSearchDT011_Click()

    m_blnTotalFlg = False
    frmSearchDT011.Show vbModal
    If m_blnTotalFlg = False Then
        imnPrice.SetFocus
    Else
        imtPnum.SetFocus
    End If
    DoEvents

End Sub

Private Sub cmdUekisearch_Click()

    frmUekiSearch.Show vbModal
    
End Sub

Private Sub cmdYpmf020Clear_Click()
    
    Dim strSQL As String
    
    If MsgBox("���̓��[�N���N���A���܂����H", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    '���͓r���̃��[�N�f�[�^�S�폜
    strSQL = "DELETE FROM YPMF020"
    g_clsAdoSQL.Connection.Execute strSQL
        
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
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
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub Form_Load()
    
    Dim strSQL As String
    
    On Error GoTo Form_Load_Err

'    Me.Caption = SYSTEM_NAME & "-" & "�������ʓ���"
    Me.Caption = "�������ʓ���"
    
    '�t�H�[���̃Z���^�����O
'    Me.left = (Screen.Width - Me.Width) / 2
'    Me.top = 0
    
    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    'PC���擾
    m_strPcName = Trim(Global_GetPcName)

    '���͓r���̃��[�N�f�[�^�폜
    strSQL = "DELETE FROM YPMF020" & _
             " WHERE PcName = '" & m_strPcName & "'"
    g_clsAdoSQL.Connection.Execute strSQL

    '�����{�^��
    optSyori(0).Value = True
    optSyori(1).Value = False
    optSyori(2).Value = False
    optSyori(3).Value = False
    optSyori(4).Value = False

    chkAutoCode.Value = AUTO_CODE
    If chkAutoCode.Value = 1 Then txtOcode.Text = AutoCodeSet()
    m_strLastOcode = ""
    
    Call PriceTani
    
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
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub Form_Unload(Cancel As Integer)

    Dim strSQL As String

    On Error GoTo Form_Unload_Err
    
    '���͓r���̃��[�N�f�[�^�폜
    strSQL = "DELETE FROM YPMF020" & _
             " WHERE PcName = '" & m_strPcName & "'"
    g_clsAdoSQL.Connection.Execute strSQL
    
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
'���@���@�@�F0�F�S��� 1:�L�[�� 2:�L�[���Ɩ��ו� 3���ו�
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub FieldsClear(intKubun As Integer)

    Dim strSQL As String

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        txtOcode.Text = ""
        txtOcode.Tag = ""
        '�w�b�_�[
        txtHnum.Text = ""
        txtHnum.Tag = ""
        '����
        imnNo.Value = 1
        imtPnum.Text = ""
        imtPnum.Tag = ""
        imtPnumLine.Text = ""
        imtPnumLine.Tag = ""
        cboIcode.Text = ""
        txtIname.Text = ""
        imnQty.Value = 0
        imnQty.Enabled = False
        imnPrice.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkWdiv.Value = 0
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
        lblDetailPcode.Caption = ""
        lblDetailPname.Caption = ""
        lblItime.Caption = ""
        imnSline.Value = 0
        chkIdiv.Value = 0
        imnOrgNum.Value = 0
        
        lsvMeisai.ListItems.Clear
        imnPrice_Total.Value = 0
    
        '���׍폜���[�N�̃N���A
        ReDim g_usrMeisaiDel(0)
    
    ElseIf intKubun = 1 Then
        txtOcode.Text = ""
        txtOcode.Tag = ""
    ElseIf intKubun = 2 Then
        '�w�b�_�[
        txtHnum.Text = ""
        txtHnum.Tag = ""
        '����
        imnNo.Value = 1
        imtPnum.Text = ""
        imtPnum.Tag = ""
        imtPnumLine.Text = ""
        imtPnumLine.Tag = ""
        cboIcode.Text = ""
        txtIname.Text = ""
        imnQty.Value = 0
        imnQty.Enabled = False
        imnPrice.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkWdiv.Value = 0
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
        lblDetailPcode.Caption = ""
        lblDetailPname.Caption = ""
        lblItime.Caption = ""
        imnSline.Value = 0
        chkIdiv.Value = 0
        imnOrgNum.Value = 0
        
        lsvMeisai.ListItems.Clear
        imnPrice_Total.Value = 0
    
        '���׍폜���[�N�̃N���A
        ReDim g_usrMeisaiDel(0)
    
    ElseIf intKubun = 3 Then
        '����
        'imnNo.Value = 1
        imtPnum.Text = ""
        imtPnum.Tag = ""
        imtPnumLine.Text = ""
        imtPnumLine.Tag = ""
        cboIcode.Text = ""
        txtIname.Text = ""
        imnQty.Value = 0
        imnPrice.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkWdiv.Value = 0
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
        lblDetailPcode.Caption = ""
        lblDetailPname.Caption = ""
        lblItime.Caption = ""
        imnSline.Value = 0
        chkIdiv.Value = 0
        imnOrgNum.Value = 0
    End If
        
    If intKubun <> 3 Then
        '���͓r���̃��[�N�f�[�^�폜
        strSQL = "DELETE FROM YPMF020" & _
                 " WHERE PcName = '" & m_strPcName & "'"
        g_clsAdoSQL.Connection.Execute strSQL
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("��ʃN���A�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

Private Sub imnBnum_GotFocus()
    
    imnBnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBnum_LostFocus()
    
    imnBnum.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnSline_GotFocus()
    
    imnSline.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnSline_LostFocus()
    
    imnSline.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnSnum_GotFocus()
    
    imnSnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnSnum_LostFocus()
    
    imnSnum.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnPrice_GotFocus()
    
    imnPrice.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnPrice_LostFocus()
    
    imnPrice.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnQty_GotFocus()
    
    imnQty.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnQty_LostFocus()
    
    imnQty.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imtIcode_Kana_Focus1_GotFocus()

'    On Error Resume Next

'    If Trim(txtIname.Text) = "" Then
        txtIname.SetFocus
'    Else
'        imnQty.SetFocus
'    End If

End Sub

Private Sub imtIcode_Kana_Focus2_GotFocus()

    imtPnum.SetFocus

End Sub

Private Sub imtPnum_Focus1_GotFocus()

    On Error GoTo imtPnum_Focus1_GotFocus_Err
    
    If Trim(imtPnum.Text) = "" Then
        imnPrice.SetFocus
        Exit Sub
    End If
    
    '�o�O�H�΍�
    If imtPnum_Focus1.Tag = "already" Then
        If m_blnTotalFlg = False Then
            imnPrice.SetFocus
        Else
            imtPnum.SetFocus
        End If
        DoEvents
        imtPnum_Focus1.Tag = ""
        Exit Sub
    End If
    
    imtPnum_Focus1.Tag = "already"
    
    '�����t�H�[���\��
    m_blnTotalFlg = False
    frmSearchDT011.Show vbModal
    If m_blnTotalFlg = False Then
        imnPrice.SetFocus
    Else
        imtPnum.SetFocus
    End If
    
    Exit Sub

imtPnum_Focus1_GotFocus_Err:

    Call MsgBox("�t�H�[�J�X�擾���G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "imtPnum_Focus1_GotFocus_Err")

End Sub

Private Sub imtPnum_Focus2_GotFocus()

    imtPnum.SetFocus

End Sub

Private Sub imtPnum_GotFocus()
    
    imtPnum.BackColor = FOCUS_STOP_COLOR
    imtPnum.Tag = imtPnum.Text
    
End Sub

Private Sub imtPnum_LostFocus()
    
    imtPnum.BackColor = FOCUS_NO_COLOR
    imtPnum.Tag = ""

End Sub

Private Sub imtPnum_Validate(Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo imtPnum_Validate_Validate_Err
    
    If Trim(imtPnum.Text) = "" Then
        imtPnumLine.Text = ""
        Exit Sub
    End If
    If imtPnum.Tag = imtPnum.Text Then Exit Sub
    If imtPnum.Tag <> imtPnum.Text Then imtPnumLine.Text = ""
    
    With adoRecordset1
        '����f�[�^
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & imtPnum.Text
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If .EOF Then
            imtPnumLine.Text = ""
            Cancel = True
            Call MsgBox("����f�[�^���o�^����Ă��܂���B", vbOKOnly + vbCritical, "")
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Exit Sub

imtPnum_Validate_Validate_Err:

    Call MsgBox("�t�H�[�J�X�ړ��O�G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "imtPnum_Validate_Validate_Err")

End Sub

Private Sub imtPnumLine_GotFocus()
    
    imtPnumLine.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imtPnumLine_LostFocus()
    
    imtPnumLine.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imtPnumLine_Validate(Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim intIndex1 As Integer

    On Error GoTo imtPnumLine_Validate_Err
    
    If Trim(imtPnumLine.Text) = "" Then Exit Sub
    If Trim(imtPnum.Text) = "" Then
        imtPnumLine.Text = ""
        Exit Sub
    End If
    
    '���ׂ���T��
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        If CLng(lsvMeisai.ListItems(intIndex1).SubItems(1)) = CLng(imtPnum.Text) Then
            If CLng(lsvMeisai.ListItems(intIndex1).SubItems(2)) = CLng(imtPnumLine.Text) Then
                Cancel = True
                Call MsgBox("���ɓo�^����Ă��܂��B", vbOKOnly + vbCritical, "")
                Exit Sub
            End If
        End If
    Next intIndex1
    
    With adoRecordset1
        '������׃f�[�^
        strSQL = "SELECT * FROM DT011" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & imtPnum.Text & _
                 " AND Line = " & imtPnumLine.Text
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If .EOF Then
            Cancel = True
            Call MsgBox("����f�[�^���o�^����Ă��܂���B", vbOKOnly + vbCritical, "")
            Exit Sub
        Else
            If Not IsNull(.Fields("Idiv")) Then
                If .Fields("Idiv") = INPUT_ON Then
                    Cancel = True
                    Call MsgBox("���ɓo�^����Ă��܂��B", vbOKOnly + vbCritical, "")
                    Exit Sub
                End If
            End If
        End If
        cboIcode.Text = IIf(IsNull(.Fields("Icode")), "", .Fields("Icode"))
        txtIname.Text = IIf(IsNull(.Fields("Iname")), "", .Fields("Iname"))
        imnQty.Value = IIf(IsNull(.Fields("Qty")), 0, .Fields("Qty"))
        
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Exit Sub

imtPnumLine_Validate_Err:

    Call MsgBox("�t�H�[�J�X�ړ��O�G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "imtPnumLine_Validate_Err")

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    On Error Resume Next

    If optSyori(0).Value = True Then
        txtHnum.SetFocus
    Else
        txtOcode.SetFocus
    End If

End Sub

Private Sub lsvMeisai_Click()

    On Error Resume Next

    '�s���I������Ă��邩�H
    If lsvMeisai.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    '���ו\��
    Call ListViewGetItem
    
    imtPnum.SetFocus

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�����敪�{�^���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub optSyori_Click(Index As Integer)

    Dim intIndex1 As Integer
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo optSyori_Click_Err

    '��ʃN���A
    Call FieldsClear(0)
    
    '�w�i�F�̕ύX
    For intIndex1 = 0 To 4
        If intIndex1 = Index Then
            optSyori(intIndex1).BackColor = BUTTON_ON
        Else
            optSyori(intIndex1).BackColor = BUTTON_OFF
        End If
    Next intIndex1
    
    Select Case Index
        Case 0: '�V�K
            Call FieldsControl(0, True)
            Call FieldsControl(1, False)
            Call FieldsControl(2, True)
            Call FieldsControl(3, True)
            If chkAutoCode.Value = 1 Then txtOcode.Text = AutoCodeSet
        Case 1: '�ύX
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
            Call FieldsControl(3, True)
        Case 2: '�폜
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
            Call FieldsControl(3, False)
        Case 3: '���
            Call FieldsControl(0, True)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
            Call FieldsControl(3, False)
            m_strLastOcode = ""
            frmPrintDialog.m_blnAutoPrint = False
            frmPrintDialog.Show vbModal
        Case 4: '�O���o��
    End Select

    On Error Resume Next
    If Index = 0 Then
        txtHnum.SetFocus
    Else
        txtOcode.SetFocus
    End If
    DoEvents
    
    Exit Sub

optSyori_Click_Err:

    Call MsgBox("�����敪�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")

End Sub

Private Sub txtIname_GotFocus()

    txtIname.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtIname_LostFocus()

    txtIname.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtOcode_Change()

    If Trim(txtOcode.Text) = "" Then Exit Sub

    If txtOcode.Tag <> txtOcode.Text Then
        If optSyori(0).Value Or optSyori(1).Value Then
            fraMeisai.Enabled = True
            DoEvents
        End If
    End If

End Sub

Private Sub txtOcode_GotFocus()

    txtOcode.Tag = txtOcode.Text
    txtOcode.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtOcode_LostFocus()

    txtOcode.Tag = ""
    txtOcode.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtOcode_Validate(Cancel As Boolean)

    If Trim(txtOcode.Text) = "" Then Exit Sub
    If txtOcode.Tag = txtOcode.Text Then Exit Sub

    If optSyori(0).Value = True Then
        If FieldsSet(False) = True Then
            Cancel = True
            Call MsgBox("���Ƀf�[�^�����݂��܂��B", vbOKOnly + vbCritical, "")
            Exit Sub
        End If
    Else
        If FieldsSet(True) = False Then
            Cancel = True
            Call MsgBox("�f�[�^�����݂��܂���B", vbOKOnly + vbCritical, "")
            Exit Sub
        End If
    End If

End Sub

'�ځ@�I�@�@�F���̓`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "�J�ÔN��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtOcode.Text) = "" Then
        strErrMsg = "�����ԍ�����͂��Ă��������B"
        txtOcode.SetFocus
        GoTo ErrorTrap:
    End If
    If lsvMeisai.ListItems.Count <= 0 Then
        strErrMsg = "���ׂ���͂��Ă��������B"
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

'�ځ@�I�@�@�F�t�B�[���h�̐���
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FintKbn 0:�L�[�� 1:���R�[�h�ړ� 2:���ׁ@3:�w�b�_�[
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub FieldsControl(intKbn As Integer, blnEnabled As Boolean)

    On Error GoTo FieldsControl_Err
    
    Select Case intKbn
        Case 0:
            fraKey.Enabled = blnEnabled
        Case 1:
            fraRecordSelector.Enabled = blnEnabled
            cmdDataMove(0).Enabled = blnEnabled
            cmdDataMove(1).Enabled = blnEnabled
            cmdDataMove(2).Enabled = blnEnabled
            cmdDataMove(3).Enabled = blnEnabled
            cmdSearch.Enabled = blnEnabled
        Case 2:
            fraMeisai.Enabled = blnEnabled
        Case 3:
            fraHeader.Enabled = blnEnabled
    End Select
        
    Exit Sub

FieldsControl_Err:

    Call MsgBox("�t�B�[���h�̐���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsControl_Err")

End Sub

'�ځ@�I�@�@�F�t�B�[���h�̃Z�b�g
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Public Function FieldsSet(blnVisible As Boolean, Optional adoRecordsetArg As Variant) As Boolean
    
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim strSQL As String
    Dim itmX As ListItem
    Dim intIndex1 As Integer

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    If IsMissing(adoRecordsetArg) = False Then
        Set adoRecordset1 = adoRecordsetArg
    Else
        strSQL = "SELECT * FROM DT020" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Ocode = '" & txtOcode.Text & "'"
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    End If
    
    With adoRecordset1
        If .EOF Or .BOF Then
            adoRecordset1.Close
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        If blnVisible = False Then
            adoRecordset1.Close
            FieldsSet = True
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        txtOcode.Text = .Fields("Ocode")
        txtHnum.Text = IIf(IsNull(.Fields("Hnum")), "", Trim(.Fields("Hnum")))
        .Close
    End With
    
    With adoRecordset2
        intIndex1 = 1
        lsvMeisai.ListItems.Clear
        
        strSQL = "SELECT * FROM DT021" & _
                 " WHERE Ocode = '" & txtOcode.Text & "'" & _
                 " ORDER BY Ocode,Line"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Set itmX = lsvMeisai.ListItems.Add(, , intIndex1, 0)
            itmX.SubItems(1) = IIf(IsNull(.Fields("Pnum")), "", .Fields("Pnum"))
            itmX.SubItems(2) = IIf(IsNull(.Fields("PnumLine")), "", Trim(.Fields("PnumLine")))
            itmX.SubItems(3) = IIf(IsNull(.Fields("Icode")), "", Trim(.Fields("Icode")))
            itmX.SubItems(4) = IIf(IsNull(.Fields("Iname")), "", Trim(.Fields("Iname")))
            itmX.SubItems(5) = IIf(IsNull(.Fields("Qty")), 0, Format(.Fields("Qty"), "#,##0"))
            itmX.SubItems(6) = IIf(IsNull(.Fields("Price")), 0, Format(.Fields("Price"), "#,##0"))
            itmX.SubItems(7) = IIf(IsNull(.Fields("Bcode")), "", Trim(.Fields("Bcode")))
            If Trim(itmX.SubItems(7)) <> "" Then
                itmX.SubItems(8) = Global_Get_Bname(g_clsAdoSQL, itmX.SubItems(7), txtOcode.Text, "")
            Else
                itmX.SubItems(8) = ""
            End If
            If Not IsNull(.Fields("Wdiv")) Then
                itmX.SubItems(9) = IIf(.Fields("Wdiv"), 1, 0)
            Else
                itmX.SubItems(9) = 0
            End If
            If Not IsNull(.Fields("Sdiv")) Then
                itmX.SubItems(10) = IIf(.Fields("Sdiv"), 1, 0)
            Else
                itmX.SubItems(10) = 0
            End If
            If Not IsNull(.Fields("Bdiv")) Then
                itmX.SubItems(11) = IIf(.Fields("Bdiv"), 1, 0)
            Else
                itmX.SubItems(11) = 0
            End If
            itmX.SubItems(12) = IIf(IsNull(.Fields("Bnum")), 0, Trim(.Fields("Bnum")))
            itmX.SubItems(13) = IIf(IsNull(.Fields("Snum")), 0, Trim(.Fields("Snum")))
            itmX.SubItems(14) = IIf(IsNull(.Fields("Itime")), "", Trim(.Fields("Itime")))
            If Not IsNull(.Fields("Pcode")) Then
                itmX.SubItems(15) = Trim(.Fields("Pcode"))
                itmX.SubItems(16) = Get_Pname(Trim(.Fields("Pcode")))
            Else
                itmX.SubItems(15) = ""
                itmX.SubItems(16) = ""
            End If
            If Not IsNull(.Fields("Sline")) Then
                itmX.SubItems(17) = .Fields("Sline")
            Else
                itmX.SubItems(17) = "0"
            End If
            If Not IsNull(.Fields("Idiv")) Then
                itmX.SubItems(18) = .Fields("Idiv")
            Else
                itmX.SubItems(18) = "0"
            End If
            If Trim(itmX.SubItems(17)) <> "" And itmX.SubItems(17) <> "0" Then
                itmX.SubItems(19) = "��"
            End If
            If Trim(itmX.SubItems(18)) <> "" And itmX.SubItems(18) <> "0" Then
                itmX.SubItems(19) = "��"
            End If
            itmX.SubItems(20) = intIndex1
            
            intIndex1 = intIndex1 + 1
            .MoveNext
        Loop
        .Close
        
        Call ListViewGetMaxRow
    End With

    Call Calc_Total

    Set adoRecordset1 = Nothing
    Set adoRecordset2 = Nothing
    
    Screen.MousePointer = vbDefault
    
    FieldsSet = True
    
    Exit Function

FieldsSet_Err:

    FieldsSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�t�B�[���h�Z�b�g�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsSet_Err")

End Function

'�ځ@�I�@�@�F�f�[�^�̓o�^
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim intIndex1 As Integer
    Dim blnSeisanFlg As Boolean                 '���Z�t���O(���Z�f�[�^�쐬�ς݂̏ꍇ True)
    Dim blnHenkouFlg As Boolean                 '�ύX�t���O
    Dim strBuff As String
    
#If DebugMode = 1 Then
    Dim clsDebugLog As New clsLogfile
    Dim strDebugMsg As String
#End If
    
    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    blnSeisanFlg = False
    
    '�V�K�̏ꍇ
    If optSyori(0).Value = True Then
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            '��t�f�[�^���`�F�b�N����
            strSQL = "SELECT * FROM DT011" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Pnum = " & lsvMeisai.ListItems(intIndex1).SubItems(1) & _
                     " AND Line = " & lsvMeisai.ListItems(intIndex1).SubItems(2)
            adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoRecordset1.EOF = False Then
                If Not IsNull(adoRecordset1.Fields("Idiv")) Then
                    If adoRecordset1.Fields("Idiv") = INPUT_ON Then
                        strBuff = "�s�ԍ�:" & CStr(intIndex1) & vbCrLf & _
                                  "��t�ԍ�:" & lsvMeisai.ListItems(intIndex1).SubItems(1) & vbCrLf & _
                                  "��t�s�ԍ�:" & lsvMeisai.ListItems(intIndex1).SubItems(2) & vbCrLf & _
                                  "���ɓ��͂���Ă��܂�����낵���ł����H"
                        If MsgBox(strBuff, vbInformation + vbYesNo, "") = vbNo Then
                            adoRecordset1.Close
                            DataUpdate = False
                            Screen.MousePointer = vbDefault
                            Exit Function
                        End If
                    End If
                End If
            End If
            adoRecordset1.Close
        Next intIndex1
    ElseIf optSyori(1).Value = True Then
        '�ύX�̏ꍇ
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            '�o�i�Ґ��Z�敪�Ɣ��吸�Z�敪���`�F�b�N
            If lsvMeisai.ListItems(intIndex1).SubItems(10) = "1" Or lsvMeisai.ListItems(intIndex1).SubItems(11) = "1" Then
                blnSeisanFlg = True
                Exit For
            End If
        Next intIndex1
    End If
    
'********** �����J�n **********

    g_clsAdoSQL.Connection.BeginTrans
    
    If optSyori(0).Value = True Then
        '�����̔�
        txtOcode.Text = AutoCodeSet()
    
        '������׃f�[�^�̃t���O��߂�
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)) And IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(2)) Then
                If DataDelete_DT011(lsvMeisai.ListItems(intIndex1).SubItems(1), lsvMeisai.ListItems(intIndex1).SubItems(2)) = False Then Error 1
            End If
        Next intIndex1
        
    ElseIf optSyori(1).Value = True Then
        
        '������׃f�[�^�̃t���O��߂�
        strSQL = "SELECT * FROM DT021" & _
                 " WHERE Ocode = '" & txtOcode.Text & "'"
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset2.EOF
            If IsNull(adoRecordset2.Fields("Pnum")) = False And IsNull(adoRecordset2.Fields("PnumLine")) = False Then
                If DataDelete_DT011(adoRecordset2.Fields("Pnum"), adoRecordset2.Fields("PnumLine")) = False Then Error 1
            End If
            adoRecordset2.MoveNext
        Loop
        adoRecordset2.Close
        
        '���Z�f�[�^�쐬�ς݂̏ꍇ
        If blnSeisanFlg = True Then
            '�ύX�f�[�^�̃`�F�b�N
            blnHenkouFlg = False
            If DataUpdate_CheckData(blnHenkouFlg) = False Then Error 1
        End If
    End If
 
    '�f�[�^�폜
    strSQL = "DELETE FROM DT021" & _
             " WHERE Ocode = '" & txtOcode.Text & "'"
    g_clsAdoSQL.Connection.Execute strSQL
    
    strSQL = "DELETE FROM DT020" & _
             " WHERE Ocode = '" & txtOcode.Text & "'"
    g_clsAdoSQL.Connection.Execute strSQL
 
    With adoRecordset1
        strSQL = "SELECT * FROM DT020" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Ocode = '" & txtOcode.Text & "'"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        .Fields("Odate") = lblOdate.Caption
        .Fields("Ocode") = txtOcode.Text
        .Fields("Hnum") = txtHnum.Text
        .Update
        .Close
        
#If DebugMode = 1 Then
    '�f�o�b�O���O���o��
    strDebugMsg = Format$(Now(), "yyyy/mm/dd hh:mm:ss") & "-->"
    If optSyori(0).Value = True Then
        strDebugMsg = strDebugMsg & "Ocode=" & txtOcode.Text & "�@���O�J�n ---- �������[�h(�V�K)"
    ElseIf optSyori(1).Value = True Then
        strDebugMsg = strDebugMsg & "Ocode=" & txtOcode.Text & "�@���O�J�n ---- �������[�h(�ύX)"
    End If
    clsDebugLog.SetMessage App.Path & "\" & "Ypmf020_" & Format$(Now(), "yyyymmdd") & ".log", "Ypmf020", "DataUpdate", strDebugMsg, 0
#End If
    
    End With
    
    With adoRecordset2
        strSQL = "SELECT * FROM DT021"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            .AddNew
            .Fields("Ocode") = txtOcode.Text
            .Fields("Line") = intIndex1
            .Fields("Pnum") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)), lsvMeisai.ListItems(intIndex1).SubItems(1), Null)
            .Fields("PnumLine") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(2)), lsvMeisai.ListItems(intIndex1).SubItems(2), Null)
            .Fields("Icode") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(3)), lsvMeisai.ListItems(intIndex1).SubItems(3), Null)
            .Fields("Iname") = lsvMeisai.ListItems(intIndex1).SubItems(4)
            .Fields("Qty") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(5)), lsvMeisai.ListItems(intIndex1).SubItems(5), 0)
            .Fields("Price") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(6)), lsvMeisai.ListItems(intIndex1).SubItems(6), 0)
            .Fields("Bcode") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(7)), lsvMeisai.ListItems(intIndex1).SubItems(7), Null)
            .Fields("Wdiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(9)), lsvMeisai.ListItems(intIndex1).SubItems(9), 0)
            .Fields("Sdiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(10)), lsvMeisai.ListItems(intIndex1).SubItems(10), 0)
            .Fields("Bdiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(11)), lsvMeisai.ListItems(intIndex1).SubItems(11), 0)
            .Fields("Bnum") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(12)), lsvMeisai.ListItems(intIndex1).SubItems(12), 0)
            .Fields("Snum") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(13)), lsvMeisai.ListItems(intIndex1).SubItems(13), 0)
            .Fields("Sline") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(17)), lsvMeisai.ListItems(intIndex1).SubItems(17), 0)
            .Fields("Idiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(18)), lsvMeisai.ListItems(intIndex1).SubItems(18), AUCTION_ON)
            .Fields("Itime") = Format(Now(), "yyyy/mm/dd hh:mm:ss")
            .Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
            .Update
            
            '������׃f�[�^�̍X�V
            If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)) And IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(2)) Then
                If DataUpdate_DT011(lsvMeisai.ListItems(intIndex1).SubItems(1), lsvMeisai.ListItems(intIndex1).SubItems(2)) = False Then Error 1
            End If
            
#If DebugMode = 1 Then
    '�f�o�b�O���O���o��
    strDebugMsg = Format$(Now(), "yyyy/mm/dd hh:mm:ss") & "-->"
    strDebugMsg = strDebugMsg & "Ocode=" & txtOcode.Text & ","
    strDebugMsg = strDebugMsg & "Pnum=" & IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)), lsvMeisai.ListItems(intIndex1).SubItems(1), "") & ","
    strDebugMsg = strDebugMsg & "PnumLine=" & IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(2)), lsvMeisai.ListItems(intIndex1).SubItems(2), "") & ","
    strDebugMsg = strDebugMsg & "Iname=" & lsvMeisai.ListItems(intIndex1).SubItems(4) & ","
    strDebugMsg = strDebugMsg & "Bcode=" & IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(7)), lsvMeisai.ListItems(intIndex1).SubItems(7), "") & ","
    strDebugMsg = strDebugMsg & "Sdiv=" & IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(10)), lsvMeisai.ListItems(intIndex1).SubItems(10), "") & ","
    strDebugMsg = strDebugMsg & "Bdiv=" & IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(11)), lsvMeisai.ListItems(intIndex1).SubItems(11), "")
    
    clsDebugLog.SetMessage App.Path & "\" & "Ypmf020_" & Format$(Now(), "yyyymmdd") & ".log", "Ypmf020", "DataUpdate", strDebugMsg, 0
#End If
            
        Next intIndex1
        .Close
    End With
    
    g_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    Set adoRecordset1 = Nothing
    Set adoRecordset2 = Nothing
    
    If blnHenkouFlg = True Then
        Call MsgBox("���Z����Ă���f�[�^��ύX���܂����B" & vbCrLf & "�ύX�O�ƕύX��̔���E�o�i�҂̓`�[���o�͂��Ă��������B", vbOKOnly + vbInformation, "")
    End If
    
    DataUpdate = True
    
    Exit Function

DataUpdate_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    DataUpdate = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�f�[�^�o�^�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_Err")

End Function

'�ځ@�I�@�@�F�f�[�^�̍폜
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataDelete() As Boolean

    Dim strSQL As String
    Dim intIndex1 As Integer
    Dim adoRecordset1 As New ADODB.Recordset

#If DebugMode = 1 Then
    Dim clsDebugLog As New clsLogfile
    Dim strDebugMsg As String
#End If
    
    On Error GoTo DataDelete_Err
    
    Screen.MousePointer = vbHourglass
    
'    For intIndex1 = 1 To lsvMeisai.ListItems.Count
'        If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(7)) = True Then
'            '���吸�Z�f�[�^
'            strSQL = "SELECT * FROM DT041" & _
'                     " WHERE Odate = '" & Trim(lblOdate.Caption) & "'" & _
'                     " AND Bcode = " & lsvMeisai.ListItems(intIndex1).SubItems(7) & _
'                     " ORDER BY Odate,Bcode,Num"
'            adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
'            If adoRecordset1.EOF = False Then
'                DataDelete = False
'                Screen.MousePointer = vbDefault
'                Call MsgBox("���ɐ��Z����Ă��邽�ߍ폜�ł��܂���B", vbOKOnly + vbCritical, "")
'                Exit Function
'            End If
'            adoRecordset1.Close
'        End If
'    Next intIndex1
    
    With g_clsAdoSQL.Connection
        .BeginTrans
        
#If DebugMode = 1 Then
    '�f�o�b�O���O���o��
    strDebugMsg = Format$(Now(), "yyyy/mm/dd hh:mm:ss") & "-->"
    strDebugMsg = strDebugMsg & "Ocode=" & txtOcode.Text & "�@���O�J�n"
    clsDebugLog.SetMessage App.Path & "\" & "Ypmf020_" & Format$(Now(), "yyyymmdd") & ".log", "Ypmf020", "DataDelete", strDebugMsg, 0
#End If
        
        strSQL = "SELECT * FROM DT021" & _
                 " WHERE Ocode = '" & txtOcode.Text & "'" & _
                 " ORDER BY Ocode,Line"
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset1.EOF
            If Not IsNull(adoRecordset1.Fields("Pnum")) Then
                If DataDelete_DT040(adoRecordset1.Fields("Pnum")) = False Then Error 1
            End If
            If Not IsNull(adoRecordset1.Fields("Bcode")) Then
                If DataDelete_DT041(adoRecordset1.Fields("Bcode")) = False Then Error 1
            End If
            
#If DebugMode = 1 Then
    '�f�o�b�O���O���o��
    strDebugMsg = Format$(Now(), "yyyy/mm/dd hh:mm:ss") & "-->"
    strDebugMsg = strDebugMsg & "Ocode=" & txtOcode.Text & ","
    strDebugMsg = strDebugMsg & "Pnum=" & IIf(IsNull(adoRecordset1.Fields("Pnum")), "", adoRecordset1.Fields("Pnum")) & ","
    strDebugMsg = strDebugMsg & "PnumLine=" & IIf(IsNull(adoRecordset1.Fields("PnumLine")), "", adoRecordset1.Fields("PnumLine")) & ","
    strDebugMsg = strDebugMsg & "Iname=" & IIf(IsNull(adoRecordset1.Fields("Iname")), "", adoRecordset1.Fields("Iname")) & ","
    strDebugMsg = strDebugMsg & "Bcode=" & IIf(IsNull(adoRecordset1.Fields("Bcode")), "", adoRecordset1.Fields("Bcode")) & ","
    strDebugMsg = strDebugMsg & "Sdiv=" & IIf(IsNull(adoRecordset1.Fields("Sdiv")), "", adoRecordset1.Fields("Sdiv")) & ","
    strDebugMsg = strDebugMsg & "Bdiv=" & IIf(IsNull(adoRecordset1.Fields("Bdiv")), "", adoRecordset1.Fields("Bdiv"))
    
    clsDebugLog.SetMessage App.Path & "\" & "Ypmf020_" & Format$(Now(), "yyyymmdd") & ".log", "Ypmf020", "DataDelete", strDebugMsg, 0
#End If
            
            adoRecordset1.MoveNext
        Loop
        
        strSQL = "DELETE FROM DT021" & _
                 " WHERE Ocode = '" & txtOcode.Text & "'"
        .Execute strSQL
    
        strSQL = "DELETE FROM DT020" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Ocode = '" & txtOcode.Text & "'"
        .Execute strSQL
    
        '������׃f�[�^�̃t���O��߂�
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            If DataDelete_DT011(lsvMeisai.ListItems(intIndex1).SubItems(1), lsvMeisai.ListItems(intIndex1).SubItems(2)) = False Then Error 1
        Next intIndex1
    
        .CommitTrans
    End With
    
    Screen.MousePointer = vbDefault
    
    DataDelete = True
    
    Exit Function

DataDelete_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    DataDelete = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�f�[�^�̍폜�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataDelete_Err")

End Function

'�ځ@�I�@�@�F�R�[�h�̎����̔�
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function AutoCodeSet() As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo AutoCodeSet_Err
    
    AutoCodeSet = ""
    
    Screen.MousePointer = vbHourglass
    
    With adoRecordset1
        strSQL = "SELECT Ocode FROM DT020" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " ORDER BY Odate ASC,Ocode DESC"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If .EOF Or .BOF Then
            AutoCodeSet = CStr(Global_Get_NumericDay(Trim(lblOdate.Caption))) & "0001"
            adoRecordset1.Close
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        AutoCodeSet = left$(.Fields("Ocode"), 8) & Format(CLng(right$(.Fields("Ocode"), 4)) + 1, "0000")
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Function

AutoCodeSet_Err:

    AutoCodeSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�R�[�h�̎����̔ԃG���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "AutoCodeSet_Err")

End Function

'�ځ@�I�@�@�F���X�g�r���[�ւ̃f�[�^�o�^
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FintFlg(0:�ǉ��E�X�V 1:�}��)
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function ListViewSetItem(intPostion As Integer, intFlg As Integer) As Boolean

    Dim itmX As ListItem

    On Error GoTo ListViewSetItem_Err
    
    ListViewSetItem = False
    
    '���X�g�r���[�̃f�[�^�����i�s�ԍ�����v����f�[�^����������폜�j
    Set itmX = lsvMeisai.FindItem(intPostion, , , 0)
    If Not (itmX Is Nothing) Then
        If intFlg = 0 Then
            '�f�[�^�폜
            lsvMeisai.ListItems.Remove itmX.Index
        End If
        '�f�[�^��ǉ�
        Set itmX = lsvMeisai.ListItems.Add(intPostion, , intPostion, 0)
    Else
        '�f�[�^��ǉ�
        Set itmX = lsvMeisai.ListItems.Add(, , intPostion, 0)
    End If
    itmX.SubItems(1) = Trim(imtPnum.Text)
    itmX.SubItems(2) = Trim(imtPnumLine.Text)
    itmX.SubItems(3) = Trim(cboIcode.Text)
    itmX.SubItems(4) = Trim(txtIname.Text)
    itmX.SubItems(5) = Format(imnQty.Value, "#,##0")
    itmX.SubItems(6) = Format(CCur(imnPrice.Value) * m_curPriceTani, "#,##0")
    itmX.SubItems(7) = Trim(cboBcode.Text)
    itmX.SubItems(8) = Trim(lblBname.Caption)
    itmX.SubItems(9) = chkWdiv.Value
    itmX.SubItems(10) = chkSdiv.Value
    itmX.SubItems(11) = chkBdiv.Value
    itmX.SubItems(12) = imnBnum.Value
    itmX.SubItems(13) = imnSnum.Value
    itmX.SubItems(14) = Trim(lblItime.Caption)
    itmX.SubItems(15) = Trim(lblDetailPcode.Caption)
    itmX.SubItems(16) = Trim(lblDetailPname.Caption)
    itmX.SubItems(17) = imnSline.Value
    itmX.SubItems(18) = chkIdiv.Value
    If Trim(itmX.SubItems(17)) <> "" And itmX.SubItems(17) <> "0" Then
        itmX.SubItems(19) = "��"
    End If
    If Trim(itmX.SubItems(18)) <> "" And itmX.SubItems(18) <> "0" Then
        itmX.SubItems(19) = "��"
    End If
    itmX.SubItems(20) = imnOrgNum.Value

    '���X�g�r���[���X�N���[�����āA���o���ꂽ ListItem ��\��
    lsvMeisai.ListItems(lsvMeisai.ListItems.Count).EnsureVisible
    
    '�s�ԍ��擾
    Call ListViewGetMaxRow
    
    ListViewSetItem = True
    
    Exit Function

ListViewSetItem_Err:

    Call MsgBox("���X�g�r���[�ւ̃f�[�^�o�^�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewSetItem_Err")

End Function

'�ځ@�I�@�@�F���X�g�r���[����̃f�[�^�\��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub ListViewGetItem()

    On Error GoTo ListViewGetItem_Err
    
    imnNo.Value = lsvMeisai.SelectedItem.Text
    imtPnum.Text = lsvMeisai.SelectedItem.SubItems(1)
    imtPnumLine.Text = lsvMeisai.SelectedItem.SubItems(2)
    cboIcode.Text = lsvMeisai.SelectedItem.SubItems(3)
    txtIname.Text = lsvMeisai.SelectedItem.SubItems(4)
    imnQty.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(5)) <> "", lsvMeisai.SelectedItem.SubItems(5), 0)
    imnPrice.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(6)) <> "", CCur(lsvMeisai.SelectedItem.SubItems(6)) / m_curPriceTani, 0)
    cboBcode.Text = lsvMeisai.SelectedItem.SubItems(7)
    lblBname.Caption = lsvMeisai.SelectedItem.SubItems(8)
    chkWdiv.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(9)) <> "", lsvMeisai.SelectedItem.SubItems(9), 0)
    chkSdiv.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(10)) <> "", lsvMeisai.SelectedItem.SubItems(10), 0)
    chkBdiv.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(11)) <> "", lsvMeisai.SelectedItem.SubItems(11), 0)
    imnBnum.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(12)) <> "", lsvMeisai.SelectedItem.SubItems(12), 0)
    imnSnum.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(13)) <> "", lsvMeisai.SelectedItem.SubItems(13), 0)
    lblItime.Caption = lsvMeisai.SelectedItem.SubItems(14)
    lblDetailPcode.Caption = lsvMeisai.SelectedItem.SubItems(15)
    lblDetailPname.Caption = lsvMeisai.SelectedItem.SubItems(16)
    imnSline.Text = IIf(Trim(lsvMeisai.SelectedItem.SubItems(17)) <> "", lsvMeisai.SelectedItem.SubItems(17), 0)
    chkIdiv.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(18)) <> "", lsvMeisai.SelectedItem.SubItems(18), AUCTION_ON)
    imnOrgNum.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(20)) = False, 0, lsvMeisai.SelectedItem.SubItems(20))
        
    Exit Sub
    
ListViewGetItem_Err:

   Call MsgBox("���X�g�r���[����f�[�^�擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewGetItem_Err")

End Sub

'�ځ@�I�@�@�F���X�g�r���[����̃f�[�^�폜
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function ListViewDelItem() As Boolean

    Dim itmX As ListItem
    Dim intPostion As Integer

    On Error GoTo ListViewDelItem_Err

    ListViewDelItem = False

    If MsgBox("���ׂ��폜���܂����H", vbYesNo + vbQuestion, "") = vbNo Then Exit Function
    
    '�폜�s�̎擾
    intPostion = imnNo.Value
    
    '���X�g�r���[�̃f�[�^�����i�s�ԍ�����v����f�[�^����������폜�j
    Set itmX = lsvMeisai.FindItem(intPostion, , , 0)
    If Not (itmX Is Nothing) Then
        '�f�[�^�폜
        lsvMeisai.ListItems.Remove itmX.Index
        
'        '�s�ԍ��U�蒼��
'        Call ListViewRefresh
        
        '2005/09/22 �V�K�̏ꍇ�̂�
        If optSyori(0).Value = True Then
            '�s�ԍ��U�蒼��
            Call ListViewRefresh
        End If
    End If

    '�s�ԍ��擾
    Call ListViewGetMaxRow

    ListViewDelItem = True

    Exit Function

ListViewDelItem_Err:

    Call MsgBox("���X�g�r���[����f�[�^�폜�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewDelItem_Err")

End Function

Private Sub txtHnum_GotFocus()

    txtHnum.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtHnum_LostFocus()

    txtHnum.BackColor = FOCUS_NO_COLOR

End Sub

'�ځ@�I�@�@�F���ד��̓`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DoValidationChecks_Dst() As Boolean

    Dim strErrMsg As String
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
        
    On Error GoTo DoValidationChecks_Dst_Err

    If imnNo.Value > MAX_ROW Then
        strErrMsg = StrConv((MAX_ROW + 1), vbWide) & "�s�ȏ���͂ł��܂���B"
        imtPnum.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imtPnum.Text) = "" Then
        strErrMsg = "��t�ԍ�����͂��Ă��������B"
        imtPnum.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imtPnumLine.Text) = "" Then
        strErrMsg = "��t�s�ԍ�����͂��Ă��������B"
        imtPnum.SetFocus
        GoTo ErrorTrap:
    End If
    If imnQty.Value = 0 Then
        strErrMsg = "���ʂ���͂��Ă��������B"
        imnQty.SetFocus
        GoTo ErrorTrap:
    End If
    If chkIdiv.Value = 0 And imnPrice.Value = 0 And imnSline.Value = 0 Then
        strErrMsg = "�������z����͂��Ă��������B"
        imnPrice.SetFocus
        GoTo ErrorTrap:
    End If
    If chkIdiv.Value = 0 And Trim(cboBcode.Text) = "" Then
        strErrMsg = "����R�[�h����͂��Ă��������B"
        cboBcode.SetFocus
        GoTo ErrorTrap:
    End If
    
    '2006/01/13 �ǉ�
    If chkIdiv.Value = 1 And Trim(cboBcode.Text) <> "" Then
        strErrMsg = "�����s�����̎��͔���R�[�h�͓��͂ł��܂���B"
        cboBcode.SetFocus
        GoTo ErrorTrap:
    End If
    If chkIdiv.Value = 1 And imnPrice.Value <> 0 Then
        strErrMsg = "�����s�����̎��͋��z�͓��͂ł��܂���B"
        imnPrice.SetFocus
        GoTo ErrorTrap:
    End If
    
    
    '2005/09/16 ���͒����[�N��T���B���͒��̏ꍇ�͓��͂����Ȃ�
    strSQL = "SELECT * FROM YPMF020" & _
             " WHERE Odate = '" & g_strOdate & "'" & _
             " AND Pnum = " & imtPnum.Text & _
             " AND Line = " & imtPnumLine.Text & _
             " AND PcName <> '" & m_strPcName & "'"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        adoRecordset1.Close
        strErrMsg = "���ɑ��̐l�����͒��ł��B�ʂ̍s��I�����ĉ������B"
        cboBcode.SetFocus
        GoTo ErrorTrap:
        Exit Function
    End If
    adoRecordset1.Close
        
    DoValidationChecks_Dst = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks_Dst = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "���̓`�F�b�N")
    
    Exit Function
    
DoValidationChecks_Dst_Err:

    DoValidationChecks_Dst = False
    Call MsgBox("���̓`�F�b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Dst_Err")

End Function

'�ځ@�I�@�@�F���X�g�r���[����̍s�ԍ��擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function ListViewGetMaxRow() As Boolean

    On Error GoTo ListViewGetMaxRow_Err

    ListViewGetMaxRow = False

    '�s�ԍ��擾
    imnNo.Value = lsvMeisai.ListItems.Count + 1

    ListViewGetMaxRow = True

    Exit Function

ListViewGetMaxRow_Err:

    Call MsgBox("���X�g�r���[����̍s�ԍ��擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewGetMaxRow_Err")

End Function

'�ځ@�I�@�@�F���X�g�r���[�ւ̃f�[�^�}��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function ListViewInsItem() As Boolean
    
    Dim varRes As Variant
    Dim intPostion As Integer
    
    On Error GoTo ListViewInsItem_Err
    
    ListViewInsItem = False
    
    If lsvMeisai.ListItems.Count >= MAX_ROW Then
        Call MsgBox("����ȏ㖾�ׂ���͂ł��܂���B", vbOKOnly + vbCritical, "")
        Exit Function
    End If
    
    varRes = InputBox("�}������s�ԍ�����͂��Ă�������...", "", "")

    '���͒l���`�F�b�N
    If Trim(varRes) = "" Then
        Call MsgBox("�s�ԍ�����͂��Ă��������B", vbOKOnly + vbCritical, "")
        Exit Function
    End If
    If IsNumeric(varRes) = False Then
        Call MsgBox("�s�ԍ����s���ł��B", vbOKOnly + vbCritical, "")
        Exit Function
    End If

    If DoValidationChecks_Dst() = False Then Exit Function

    '�ҏW�s�̎擾
    intPostion = CInt(varRes)
    
    Call ListViewSetItem(intPostion, 1)

    '�s�ԍ��U�蒼��
    Call ListViewRefresh

    ListViewInsItem = True

    Exit Function

ListViewInsItem_Err:

    Call MsgBox("���X�g�r���[�ւ̃f�[�^�}���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewInsItem_Err")

End Function

'�ځ@�I�@�@�F���X�g�r���[�̍s�ԍ���U�蒼��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function ListViewRefresh() As Boolean

    Dim intIndex1 As Integer

    On Error GoTo ListViewRefresh_Err

    ListViewRefresh = False

    lsvMeisai.Refresh
    For intIndex1 = 1 To lsvMeisai.ListItems.Count Step 1
        lsvMeisai.ListItems(intIndex1).Text = intIndex1
    Next intIndex1

    ListViewRefresh = True

    Exit Function

ListViewRefresh_Err:

    Call MsgBox("���X�g�r���[�̍s�ԍ���U�蒼���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewRefresh_Err")

End Function

'�ځ@�I�@�@�F�R���{�{�b�N�X�̍쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Sub MakecboBcode(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboBcode_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        strSQL = "SELECT * FROM vw_MT071" & _
                 " WHERE (Fdiv = " & BUSINESS_DIV_BUYER & " OR Fdiv = " & BUSINESS_DIV_ALL & ")"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            If Not IsNull(.Fields("Sdate")) And Not IsNull(.Fields("Fdate")) Then
                If .Fields("Sdate") <= Trim(lblOdate.Caption) And Trim(lblOdate.Caption) <= .Fields("Fdate") Then
                    Ctrl.AddItem .Fields("Bnum") & ";" & .Fields("Bname")
                Else
                    Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
                End If
            Else
                Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
            End If
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboBcode_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("�R���{�{�b�N�X�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboBcode_Err")

End Sub

'�ځ@�I�@�@�F�f�[�^�̍X�V
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataUpdate_DT011(intPnum As Integer, intLine As Integer) As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim curQty_Total As Currency

    On Error GoTo DataUpdate_DT011_Err
    
    '�����f�[�^���琔�ʂ̍��v���擾����
'    curQty_Total = 0
'    With adoRecordset2
'        strSQL = "SELECT SUM(Qty) AS Qty_Total FROM DT021" & _
'                 " WHERE LEFT(Ocode,8) = '" & Global_Get_NumericDay(lblOdate.Caption) & "'" & _
'                 " AND Pnum = " & intPnum & _
'                 " AND PnumLine = " & intLine
'        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
'        If Not .EOF Then
'            curQty_Total = .Fields("Qty_Total")
'        End If
'        .Close
'    End With
    
    With adoRecordset1
        strSQL = "SELECT * FROM DT011" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & intPnum & _
                 " AND Line = " & intLine
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            '�����f�[�^����擾�������ʂ̍��v������f�[�^�̐��ʂ��������ꍇ�̂݃t���O���X�V����
            If Not IsNull(.Fields("Qty")) Then
'                If curQty_Total >= CCur(.Fields("Qty")) Then
                    .Fields("Idiv") = INPUT_ON
                    .Update
'                End If
            End If
        End If
        .Close
    End With
    
    Set adoRecordset1 = Nothing
    
    DataUpdate_DT011 = True
    
    Exit Function

DataUpdate_DT011_Err:

    DataUpdate_DT011 = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�f�[�^�̍X�V�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_DT011_Err")

End Function

'�ځ@�I�@�@�F���̂̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�P�O
'�X�V�����@�F
'
Private Function Get_Pname(strCode As String) As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo Get_Pname_Err
    
    Get_Pname = ""
    
    If Trim(strCode) = "" Then Exit Function
    
    With adoRecordset1
        strSQL = "{call sp_MT030;2(" & strCode & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            Get_Pname = IIf(IsNull(.Fields("Pname")), "", .Fields("Pname"))
        End If
    End With
    
    Exit Function
    
Get_Pname_Err:

    Call MsgBox("���̎擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Get_Pname_Err")

End Function

'�ځ@�I�@�@�F���v�̌v�Z
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�P�R
'�X�V�����@�F
'
Public Sub Calc_Total()

    Dim intIndex1 As Integer
    Dim curTotal As Currency

    On Error GoTo Calc_Total_Err
    
    curTotal = 0
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        curTotal = curTotal + CCur(lsvMeisai.ListItems(intIndex1).SubItems(6))
    Next intIndex1
    imnPrice_Total.Value = curTotal
    
    Exit Sub
    
Calc_Total_Err:

    Call MsgBox("���v�̌v�Z�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Total_Err")

End Sub

'�ځ@�I�@�@�F�f�[�^�̍X�V(�t���O��߂�)
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataDelete_DT011(intPnum As Integer, intLine As Integer) As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo DataDelete_DT011_Err
    
    With adoRecordset1
        strSQL = "SELECT * FROM DT011" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & intPnum & _
                 " AND Line = " & intLine
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            .Fields("Idiv") = INPUT_OFF
            .Update
        End If
        .Close
    End With
    
    Set adoRecordset1 = Nothing
    
    DataDelete_DT011 = True
    
    Exit Function

DataDelete_DT011_Err:

    DataDelete_DT011 = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�f�[�^�̍X�V�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataDelete_DT011_Err")

End Function

Private Sub PriceTani()

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo PriceTani_Err

    lblPriceTani.Caption = "�~"
    imnPrice.DisplayFormat = "###,###,##0"
    imnPrice.Format = "###,###,##0"
    imnPrice.MaxValue = 999999999
    imnPrice.MinValue = -999999999
            
    m_curPriceTani = 1
        
    With adoRecordset1
        strSQL = "SELECT * FROM MT010"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Up")) Then
                Select Case CStr(.Fields("Up"))
                    Case "10":
                        lblPriceTani.Caption = "0�~"
                        imnPrice.DisplayFormat = "###,###,#0"
                        imnPrice.Format = "###,###,#0"
                        imnPrice.MaxValue = 99999999
                        imnPrice.MinValue = -99999999
                        m_curPriceTani = 10
                    Case "100":
                        lblPriceTani.Caption = "00�~"
                        imnPrice.DisplayFormat = "###,###,0"
                        imnPrice.Format = "###,###,0"
                        imnPrice.MaxValue = 9999999
                        imnPrice.MinValue = -9999999
                        m_curPriceTani = 100
                End Select
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Exit Sub

PriceTani_Err:

    Call MsgBox("�������z�̒P�ʂ̐ݒ�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "PriceTani_Err")

End Sub

'�ځ@�I�@�@�F�o�i�Ґ��Z�f�[�^�̍폜
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataDelete_DT040(intPnum As Integer) As Boolean

    Dim strSQL As String

    On Error GoTo DataDelete_DT040_Err
    
    '�������׃f�[�^
    strSQL = "UPDATE DT021" & _
             " SET Sdiv = 0," & _
             " Snum = 0" & _
             " WHERE Ocode = '" & txtOcode.Text & "'" & _
             " AND Pnum = " & intPnum
    g_clsAdoSQL.Connection.Execute strSQL
    
    '�o�i�Ґ��Z�f�[�^
    strSQL = "DELETE FROM DT040" & _
             " WHERE Odate = '" & lblOdate.Caption & "'" & _
             " AND Pnum = " & intPnum
    g_clsAdoSQL.Connection.Execute strSQL
    
    DataDelete_DT040 = True
    
    Exit Function

DataDelete_DT040_Err:

    DataDelete_DT040 = False
    Call MsgBox("�f�[�^�̍X�V�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataDelete_DT040_Err")

End Function

'�ځ@�I�@�@�F���吸�Z�f�[�^�̍폜
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataDelete_DT041(intBcode As Integer) As Boolean

    Dim strSQL As String

    On Error GoTo DataDelete_DT041_Err
    
    '�������׃f�[�^
    strSQL = "UPDATE DT031" & _
             " SET Bdiv = 0," & _
             " Bnum = 0" & _
             " WHERE Odate = '" & lblOdate.Caption & "'" & _
             " AND Bcode = " & intBcode
    g_clsAdoSQL.Connection.Execute strSQL
    
    '�������׃f�[�^
    strSQL = "UPDATE DT021" & _
             " SET Bdiv = 0," & _
             " Bnum = 0" & _
             " WHERE LEFT(Ocode,8) = '" & left$(txtOcode.Text, 8) & "'" & _
             " AND Bcode = " & intBcode
    g_clsAdoSQL.Connection.Execute strSQL
    
    '���吸�Z�f�[�^
    strSQL = "DELETE FROM DT041" & _
             " WHERE Odate = '" & lblOdate.Caption & "'" & _
             " AND Bcode = " & intBcode
    g_clsAdoSQL.Connection.Execute strSQL
    
    DataDelete_DT041 = True
    
    Exit Function

DataDelete_DT041_Err:

    DataDelete_DT041 = False
    Call MsgBox("�f�[�^�̍X�V�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataDelete_DT041_Err")

End Function

'�ځ@�I�@�@�F�ύX�f�[�^�̃`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataUpdate_CheckData(blnArgHenkouFlg As Boolean) As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim intIndex1 As Integer
    Dim blnDeleteFlg As Boolean                 '�폜�t���O
    Dim blnHenkouFlg_Bcode As Boolean           '�ύX�t���O(����)
    Dim blnHenkouFlg_Pnum As Boolean            '�ύX�t���O(��t�ԍ�)
    Dim blnHenkouFlg_Kingaku As Boolean         '�ύX�t���O(���z)
    Dim blnFlg As Boolean

    On Error GoTo DataUpdate_CheckData_Err
    
    '�������׃f�[�^
    strSQL = "SELECT * FROM DT021" & _
             " WHERE Ocode = '" & txtOcode.Text & "'" & _
             " ORDER BY Ocode,Line"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    
    Do While Not adoRecordset1.EOF
        blnDeleteFlg = True
        blnHenkouFlg_Bcode = False
        blnHenkouFlg_Pnum = False
        blnHenkouFlg_Kingaku = False
        
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            '�ύX�O�ԍ��Ɣ�r
            If Trim(lsvMeisai.ListItems(intIndex1).SubItems(20)) = Trim(adoRecordset1.Fields("Line")) Then
                '����R�[�h
                If Not IsNull(adoRecordset1.Fields("Bcode")) Then
                    If Trim(lsvMeisai.ListItems(intIndex1).SubItems(7)) <> Trim(adoRecordset1.Fields("Bcode")) Then
                        blnHenkouFlg_Bcode = True
                    End If
                Else
                    If Trim(lsvMeisai.ListItems(intIndex1).SubItems(7)) <> "" Then
                        blnHenkouFlg_Bcode = True
                    End If
                End If
                '��t�ԍ�
                If Not IsNull(adoRecordset1.Fields("Pnum")) Then
                    If Trim(lsvMeisai.ListItems(intIndex1).SubItems(1)) <> Trim(adoRecordset1.Fields("Pnum")) Then
                        blnHenkouFlg_Pnum = True
                    End If
                    '��t�s�ԍ�
                    If Trim(lsvMeisai.ListItems(intIndex1).SubItems(2)) <> Trim(adoRecordset1.Fields("PnumLine")) Then
                        blnHenkouFlg_Pnum = True
                    End If
                End If
                '���z
                If Not IsNull(adoRecordset1.Fields("Price")) Then
                    If CCur(lsvMeisai.ListItems(intIndex1).SubItems(6)) <> CCur(adoRecordset1.Fields("Price")) Then
                        blnHenkouFlg_Kingaku = True
                    End If
                Else
                    If CCur(lsvMeisai.ListItems(intIndex1).SubItems(6)) <> 0 Then
                        blnHenkouFlg_Kingaku = True
                    End If
                End If
                    
                '********** �ύX������ꍇ **********
                
                If blnHenkouFlg_Bcode = True Or blnHenkouFlg_Pnum = True Or blnHenkouFlg_Kingaku = True Then
                    
                    '********** ���[�N�̃t���O������ **********
                    
                    '���z�ύX��
                    If blnHenkouFlg_Kingaku = True Then
                        '�ύX�O�f�[�^
                        If Not IsNull(adoRecordset1.Fields("Bcode")) Then
                            If DataUpdate_CheckData_WorkUpdate1(adoRecordset1.Fields("Bcode")) = False Then Error 1
                        End If
                        If Not IsNull(adoRecordset1.Fields("Pnum")) Then
                            If DataUpdate_CheckData_WorkUpdate2(adoRecordset1.Fields("Pnum")) = False Then Error 1
                        End If
                        '�ύX��f�[�^
                        If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(7)) = True Then
                            If DataUpdate_CheckData_WorkUpdate1(lsvMeisai.ListItems(intIndex1).SubItems(7)) = False Then Error 1
                        End If
                        If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)) = True Then
                            If DataUpdate_CheckData_WorkUpdate2(lsvMeisai.ListItems(intIndex1).SubItems(1)) = False Then Error 1
                        End If
                    Else
                        '����R�[�h�ύX��
                        If blnHenkouFlg_Bcode = True Then
                            '�ύX�O�f�[�^
                            If Not IsNull(adoRecordset1.Fields("Bcode")) Then
                                If DataUpdate_CheckData_WorkUpdate1(adoRecordset1.Fields("Bcode")) = False Then Error 1
                            End If
                            '�ύX��f�[�^
                            If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(7)) = True Then
                                If DataUpdate_CheckData_WorkUpdate1(lsvMeisai.ListItems(intIndex1).SubItems(7)) = False Then Error 1
                            End If
                        End If
                        '��t�ԍ��ύX��
                        If blnHenkouFlg_Pnum = True Then
                            '�ύX�O�f�[�^
                            If Not IsNull(adoRecordset1.Fields("Pnum")) Then
                                If DataUpdate_CheckData_WorkUpdate2(adoRecordset1.Fields("Pnum")) = False Then Error 1
                            End If
                            '�ύX��f�[�^
                            If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)) = True Then
                                If DataUpdate_CheckData_WorkUpdate2(lsvMeisai.ListItems(intIndex1).SubItems(1)) = False Then Error 1
                            End If
                        End If
                    End If
                                        
                    '********** �������׃f�[�^�̃t���O������ **********
                    
                    If Not IsNull(adoRecordset1.Fields("Pnum")) Then
                        If DataDelete_DT040(adoRecordset1.Fields("Pnum")) = False Then Error 1
                    End If
                    If Not IsNull(adoRecordset1.Fields("Bcode")) Then
                        If DataDelete_DT041(adoRecordset1.Fields("Bcode")) = False Then Error 1
                    End If
                    If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)) = True Then
                        If DataDelete_DT040(lsvMeisai.ListItems(intIndex1).SubItems(1)) = False Then Error 1
                    End If
                    If IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(7)) = True Then
                        If DataDelete_DT041(lsvMeisai.ListItems(intIndex1).SubItems(7)) = False Then Error 1
                    End If
                    
                    blnArgHenkouFlg = True
                End If
                
                blnDeleteFlg = False
                Exit For
            End If
        Next intIndex1
    
        '�폜����Ă����ꍇ
        If blnDeleteFlg = True Then
            '�t���O�X�V
            If Not IsNull(adoRecordset1.Fields("Pnum")) Then
                If DataUpdate_CheckData_WorkUpdate2(adoRecordset1.Fields("Pnum")) = False Then Error 1
                If DataDelete_DT040(adoRecordset1.Fields("Pnum")) = False Then Error 1
            End If
            If Not IsNull(adoRecordset1.Fields("Bcode")) Then
                If DataUpdate_CheckData_WorkUpdate1(adoRecordset1.Fields("Bcode")) = False Then Error 1
                If DataDelete_DT041(adoRecordset1.Fields("Bcode")) = False Then Error 1
            End If
            blnArgHenkouFlg = True
        End If
    
        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close
    
    DataUpdate_CheckData = True
    
    Exit Function

DataUpdate_CheckData_Err:

    DataUpdate_CheckData = False
    Call MsgBox("�ύX�f�[�^�̃`�F�b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_CheckData_Err")

End Function

'�ځ@�I�@�@�F���[�N�̃t���O�X�V(����)
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataUpdate_CheckData_WorkUpdate1(intBode As Integer) As Boolean

    Dim intIndex1 As Integer

    On Error GoTo DataUpdate_CheckData_WorkUpdate1_Err
    
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        If Trim(lsvMeisai.ListItems(intIndex1).SubItems(7)) = CStr(intBode) Then
            lsvMeisai.ListItems(intIndex1).SubItems(11) = "0"   '���吸�Z�敪
            lsvMeisai.ListItems(intIndex1).SubItems(12) = "0"   '���吸�Z��
        End If
    Next intIndex1
    
    DataUpdate_CheckData_WorkUpdate1 = True
    
    Exit Function

DataUpdate_CheckData_WorkUpdate1_Err:

    DataUpdate_CheckData_WorkUpdate1 = False
    Call MsgBox("���[�N�̃t���O�X�V(����)�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_CheckData_WorkUpdate1_Err")

End Function

'�ځ@�I�@�@�F���[�N�̃t���O�X�V(��t)
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�P
'�X�V�����@�F
'
Private Function DataUpdate_CheckData_WorkUpdate2(intPnum As Integer) As Boolean

    Dim intIndex1 As Integer

    On Error GoTo DataUpdate_CheckData_WorkUpdate2_Err
    
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        If Trim(lsvMeisai.ListItems(intIndex1).SubItems(1)) = CStr(intPnum) Then
            lsvMeisai.ListItems(intIndex1).SubItems(10) = "0"   '�o�i�Ґ��Z�敪
            lsvMeisai.ListItems(intIndex1).SubItems(13) = "0"   '�o�i�Ґ��Z��
        End If
    Next intIndex1
    
    DataUpdate_CheckData_WorkUpdate2 = True
    
    Exit Function

DataUpdate_CheckData_WorkUpdate2_Err:

    DataUpdate_CheckData_WorkUpdate2 = False
    Call MsgBox("���[�N�̃t���O�X�V(��t)�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_CheckData_WorkUpdate2_Err")

End Function


