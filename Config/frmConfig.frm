VERSION 5.00
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   8820
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8475
   Icon            =   "frmConfig.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   8475
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame Frame4 
      Caption         =   "�o�b�N�A�b�v���"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   35
      Top             =   5340
      Width           =   8175
      Begin VB.ComboBox cboBackupDrive 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2340
         TabIndex        =   14
         Text            =   "cboBackupDrive"
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "�o�b�N�A�b�v�h���C�u"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   2115
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�z�X�g���"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   33
      Top             =   6420
      Width           =   8175
      Begin imText6Ctl.imText imtDownLoad 
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":0CFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":0D68
         Key             =   "frmConfig.frx":0D86
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin CSCmdLibCtl.CSCmdBtn cmdFile 
         Height          =   375
         Index           =   3
         Left            =   7680
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
         _Version        =   262145
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SizePicture     =   -1  'True
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   6
         rPic.top        =   6
         rPic.right      =   13
         rPic.bottom     =   13
         rText.left      =   -100
         rText.top       =   -144
         rText.right     =   256
         rText.bottom    =   187
         Picture         =   "frmConfig.frx":0DBA
      End
      Begin VB.Label Label1 
         Caption         =   "�_�E�����[�h�p�X"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�N���C�A���g���"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   8175
      Begin imText6Ctl.imText imtBin 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":120C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":127A
         Key             =   "frmConfig.frx":1298
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin imText6Ctl.imText imtLDatabase 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":12CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":133A
         Key             =   "frmConfig.frx":1358
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin imText6Ctl.imText imtLDBName 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   1440
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":138C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":13FA
         Key             =   "frmConfig.frx":1418
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin imText6Ctl.imText imtLog 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":144C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":14BA
         Key             =   "frmConfig.frx":14D8
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin CSCmdLibCtl.CSCmdBtn cmdFile 
         Height          =   375
         Index           =   0
         Left            =   7680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
         _Version        =   262145
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SizePicture     =   -1  'True
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   6
         rPic.top        =   6
         rPic.right      =   13
         rPic.bottom     =   13
         rText.left      =   -98
         rText.top       =   -142
         rText.right     =   254
         rText.bottom    =   185
         Picture         =   "frmConfig.frx":150C
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdFile 
         Height          =   375
         Index           =   1
         Left            =   7680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
         _Version        =   262145
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SizePicture     =   -1  'True
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   6
         rPic.top        =   6
         rPic.right      =   13
         rPic.bottom     =   13
         rText.left      =   -100
         rText.top       =   -144
         rText.right     =   256
         rText.bottom    =   187
         Picture         =   "frmConfig.frx":195E
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdFile 
         Height          =   375
         Index           =   2
         Left            =   7680
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
         _Version        =   262145
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SizePicture     =   -1  'True
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   6
         rPic.top        =   6
         rPic.right      =   13
         rPic.bottom     =   13
         rText.left      =   -102
         rText.top       =   -146
         rText.right     =   258
         rText.bottom    =   189
         Picture         =   "frmConfig.frx":1DB0
      End
      Begin VB.Label Label1 
         Caption         =   "���O�p�X"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   32
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "���s�t�@�C���p�X"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "���[�N�c�a�p�X"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "���[�N�c�a��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   29
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SQL�T�[�o�ڑ����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   8175
      Begin imText6Ctl.imText imtServer 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":2202
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":2270
         Key             =   "frmConfig.frx":228E
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin imText6Ctl.imText imtSQLDBName 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   960
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":22C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":2330
         Key             =   "frmConfig.frx":234E
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin imText6Ctl.imText imtDSN 
         Height          =   375
         Left            =   5280
         TabIndex        =   13
         Top             =   1920
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":2382
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":23F0
         Key             =   "frmConfig.frx":240E
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin imText6Ctl.imText imtUID 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   1440
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":2442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":24B0
         Key             =   "frmConfig.frx":24CE
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin imText6Ctl.imText imtPWD 
         Height          =   375
         Left            =   6000
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "frmConfig.frx":2502
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":2570
         Key             =   "frmConfig.frx":258E
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
         AutoConvert     =   0
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   -1
         Text            =   ""
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
      Begin imNumber6Ctl.imNumber imnQueryTimeout 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         Calculator      =   "frmConfig.frx":25C2
         Caption         =   "frmConfig.frx":25E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConfig.frx":2650
         Keys            =   "frmConfig.frx":266E
         Spin            =   "frmConfig.frx":26B8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,###,##0;-##,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,###,##0;-##,###,###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999
         MinValue        =   -99999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2011496453
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label1 
         Caption         =   "���s�^�C���A�E�g"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�p�X���[�h"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "���[�U�[��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "�f�[�^�\�[�X��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   24
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "�f�[�^�x�[�X��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "�T�[�o�[��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
   End
   Begin CSCmdLibCtl.CSCmdBtn cmdExit 
      Height          =   615
      Left            =   6660
      TabIndex        =   18
      Top             =   7680
      Width           =   1695
      _Version        =   262145
      _ExtentX        =   2990
      _ExtentY        =   1085
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
      OwnPicDisabled  =   0   'False
      CaptionPosition =   3
      rPic.left       =   7
      rPic.top        =   4
      rPic.right      =   32
      rPic.bottom     =   32
      rText.left      =   42
      rText.top       =   12
      rText.right     =   108
      rText.bottom    =   31
      Picture         =   "frmConfig.frx":26E0
   End
   Begin ComctlLib.StatusBar sbrStatusBar1 
      Align           =   2  '������
      Height          =   435
      Left            =   0
      TabIndex        =   20
      Top             =   8385
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   767
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8070
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2699
            MinWidth        =   1764
            TextSave        =   "2002/07/22"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   1252
            TextSave        =   "17:39"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1147
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1238
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   11280
      TabIndex        =   0
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmConfig.frx":283A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmConfig.frx":28A8
      Key             =   "frmConfig.frx":28C6
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
      Left            =   11520
      TabIndex        =   19
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmConfig.frx":28FA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmConfig.frx":2968
      Key             =   "frmConfig.frx":2986
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
   Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
      Height          =   615
      Left            =   4860
      TabIndex        =   17
      Top             =   7680
      Width           =   1695
      _Version        =   262145
      _ExtentX        =   2990
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "�o�^(F12)"
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
      rPic.left       =   4
      rPic.top        =   4
      rPic.right      =   32
      rPic.bottom     =   32
      rText.left      =   36
      rText.top       =   12
      rText.right     =   111
      rText.bottom    =   31
      Picture         =   "frmConfig.frx":29BA
   End
   Begin VB.Line Line2 
      X1              =   -180
      X2              =   14700
      Y1              =   7560
      Y2              =   7560
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************************
'
'   �v���O�������F�N���C�A���g���ݒ�
'   �������e�@�@�F
'   �O������@�@�F
'   �쐬�ҁ@�@�@�F������� �R���E�G���W�j�A�����O�@����
'   �쐬�N�����@�F�Q�O�O�Q�^�O�T�^�Q�S
'   �X�V�����@�@�F
'
'******************************************************************

Private m_clsReg As New clsReg          '���W�X�g������N���X

Private Declare Function GetLogicalDrives Lib "KERNEL32" () As Long

Private Sub cboBackupDrive_DropDown()

    Dim Drives As Long
    Dim I As Long
    Dim Bit As Long

    On Error GoTo cboBackupDrive_DropDown_Err

    Bit = 1
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
'���@���@�@�F�o�^�{�^���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If MsgBox("�o�^���܂����H", vbQuestion + vbYesNo, "�m�F") = vbNo Then Exit Sub

    '���̓`�F�b�N
    If DoValidationChecks() = False Then Exit Sub

    '�f�[�^�o�^
    If DataUpdate() = False Then Exit Sub
    
    '�I��
    Unload Me
    End
    
    Exit Sub

cmdExecute_Click_Err:

    Call MsgBox("�o�^�{�^���N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�o�^�{�^���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub cmdExecute_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�I���{�^���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�I���{�^���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub cmdExit_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H���_�I���{�^���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub cmdFile_Click(Index As Integer)
    
    Dim strPath As String
    
    On Error GoTo cmdFile_Click_Err
    
    '�t�H���_�I���_�C�A���O�\��
    strPath = OpenSelectFolderDialog(Me.hwnd)
    If strPath = "" Then Exit Sub
    
    Select Case Index
        Case 0:
            imtBin.Text = strPath
            imtBin.SetFocus
        Case 1:
            imtLDatabase.Text = strPath
            imtLDatabase.SetFocus
        Case 2:
            imtLog.Text = strPath
            imtLog.SetFocus
        Case 3:
            imtDownLoad.Text = strPath
            imtDownLoad.SetFocus
    End Select

    DoEvents

    Exit Sub
    
cmdFile_Click_Err:

    Call MsgBox("�t�H���_�I���{�^���N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdFile_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[�����[�h��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "���ݒ�"

    '�d���N���̃`�F�b�N
    If App.PrevInstance = True Then
        Unload Me
        End
    End If

    '��ʃN���A
    Call FildsClear
    
    '�f�[�^�\��
    Call FildsSet
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("�t�H�[�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    End

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�^�C���A�E�g�t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imnQueryTimeout_GotFocus()

    imnQueryTimeout.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�^�C���A�E�g�L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imnQueryTimeout_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�^�C���A�E�g�t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imnQueryTimeout_LostFocus()

    imnQueryTimeout.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�t�@�C���p�X�t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtBin_GotFocus()

    imtBin.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�t�@�C���p�X�L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtBin_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�t�@�C���p�X�t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtBin_LostFocus()

    imtBin.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�_�E�����[�h�p�X�t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtDownLoad_GotFocus()
    
    imtDownLoad.BackColor = FOCUS_STOP_COLOR

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�_�E�����[�h�p�X�L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtDownLoad_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�_�E�����[�h�p�X�t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtDownLoad_LostFocus()
   
    imtDownLoad.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�f�[�^�\�[�X���t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtDSN_GotFocus()

    imtDSN.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�f�[�^�\�[�X���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtDSN_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�f�[�^�\�[�X���t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtDSN_LostFocus()

    imtDSN.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[�J�X�R���g���[��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[�J�X�R���g���[��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtFocusFirst_GotFocus()

    imtBin.SetFocus

End Sub

'�ځ@�I�@�@�F�L�[�_�E�����̏���
'���@���@�@�F�e�R���g���[���̃L�[�_�E�����ɐݒ�
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Sub KeyDownControl(KeyCode As Integer, Shift As Integer)

    On Error GoTo KeyDownControl_Err
    
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

KeyDownControl_Err:

    Call MsgBox("�L�[�_�E�����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "KeyDownControl_Err")

End Sub

'�ځ@�I�@�@�F��ʃN���A
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub FildsClear()

    On Error GoTo FildsClear_Err

    imtBin.Text = ""
    imtLDatabase.Text = ""
    imtLDBName.Text = ""
    imtLog.Text = ""
    imtServer.Text = ""
    imtSQLDBName.Text = ""
    imtUID.Text = ""
    imtPWD.Text = ""
    imnQueryTimeout.Value = 0
    imtDSN.Text = ""
    cboBackupDrive.Text = ""
    imtDownLoad.Text = ""
    
    Exit Sub

FildsClear_Err:

    Call MsgBox("��ʃN���A�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FildsClear_Err")

End Sub

'�ځ@�I�@�@�F�f�[�^�\��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Public Function FildsSet() As Boolean

    On Error GoTo FildsSet_Err

    FildsSet = False

    '���W�X�g���ǂݍ���
    m_clsReg.RegKey = REG_KEY
    If m_clsReg.ReadReg = False Then Exit Function
    imtBin.Text = m_clsReg.Bin
    imtLDatabase.Text = m_clsReg.LDatabase
    imtLDBName.Text = m_clsReg.LDBName
    imtLog.Text = m_clsReg.Log
    imtServer.Text = m_clsReg.Server
    imtSQLDBName.Text = m_clsReg.DBName
    imtUID.Text = m_clsReg.UID
    imtPWD.Text = m_clsReg.PWD
    imnQueryTimeout.Text = m_clsReg.CommandTimeOut
    imtDSN.Text = m_clsReg.DSN
    cboBackupDrive.Text = m_clsReg.BackUpDrive
    imtDownLoad.Text = m_clsReg.DownLoadPath
    
    FildsSet = True
    
    Exit Function

FildsSet_Err:

    FildsSet = False
    Call MsgBox("�f�[�^�\���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FildsSet_Err")

End Function

'�ځ@�I�@�@�F���̓`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
    
    On Error GoTo DoValidationChecks_Err

    '���s�t�@�C���p�X
    If Trim(imtBin.Text) = "" Then
        imtBin.SetFocus
        strErrMsg = "���s�t�@�C���p�X����͂��Ă��������I�I"
        GoTo ErrorTrap:
    End If
    '���[�N�c�a�p�X
    If Trim(imtLDatabase.Text) = "" Then
        imtLDatabase.SetFocus
        strErrMsg = "���[�N�c�a�p�X����͂��Ă��������I�I"
        GoTo ErrorTrap:
    End If
    '���[�N�c�a��
    If Trim(imtLDBName.Text) = "" Then
        imtLDBName.SetFocus
        strErrMsg = "���[�N�c�a������͂��Ă��������I�I"
        GoTo ErrorTrap:
    End If
    '���O�p�X
    If Trim(imtLog.Text) = "" Then
        imtLog.SetFocus
        strErrMsg = "���O�p�X����͂��Ă��������I�I"
        GoTo ErrorTrap:
    End If
    '�T�[�o�[��
    If Trim(imtServer.Text) = "" Then
        imtServer.SetFocus
        strErrMsg = "�T�[�o�[������͂��Ă��������I�I"
        GoTo ErrorTrap:
    End If
    '�f�[�^�x�[�X��
    If Trim(imtSQLDBName.Text) = "" Then
        imtSQLDBName.SetFocus
        strErrMsg = "�f�[�^�x�[�X������͂��Ă��������I�I"
        GoTo ErrorTrap:
    End If
    '���[�U�[��
    If Trim(imtUID.Text) = "" Then
        imtUID.SetFocus
        strErrMsg = "���[�U�[������͂��Ă��������I�I"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks = False
    Call MsgBox(strErrMsg, vbOKOnly + vbCritical, "���̓`�F�b�N")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("���̓`�F�b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function

'�ځ@�I�@�@�F�f�[�^�o�^
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Function DataUpdate() As Boolean

    On Error GoTo DataUpdate_Err

    '�}�E�X�|�C���^�������v�ɕύX
    Screen.MousePointer = vbHourglass

    m_clsReg.Bin = imtBin.Text
    m_clsReg.LDatabase = imtLDatabase.Text
    m_clsReg.LDBName = imtLDBName.Text
    m_clsReg.Log = imtLog.Text
    m_clsReg.Server = imtServer.Text
    m_clsReg.DBName = imtSQLDBName.Text
    m_clsReg.UID = imtUID.Text
    m_clsReg.PWD = imtPWD.Text
    m_clsReg.CommandTimeOut = imnQueryTimeout.Text
    m_clsReg.DSN = imtDSN.Text
    m_clsReg.BackUpDrive = Trim(cboBackupDrive.Text)
    m_clsReg.DownLoadPath = imtDownLoad.Text
    m_clsReg.RegKey = REG_KEY
    m_clsReg.WriteReg
    
    '�}�E�X�|�C���^�����ɖ߂�
    Screen.MousePointer = vbDefault
    
    DataUpdate = True
    
    Exit Function

DataUpdate_Err:

    '�}�E�X�|�C���^�����ɖ߂�
    Screen.MousePointer = vbDefault
    DataUpdate = False
    Call MsgBox("�f�[�^�o�^�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_Err")

End Function

'�ځ@�I�@�@�F
'���@���@�@�F���[�N�c�a�p�X�t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLDatabase_GotFocus()

    imtLDatabase.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���[�N�c�a�p�X�L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLDatabase_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���[�N�c�a�p�X�t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLDatabase_LostFocus()

    imtLDatabase.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���[�N�c�a���t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLDBName_GotFocus()
    
    imtLDBName.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���[�N�c�a���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLDBName_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���[�N�c�a���t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLDBName_LostFocus()

    imtLDBName.BackColor = FOCUS_NO_COLOR

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���O�p�X�t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLog_GotFocus()

    imtLog.BackColor = FOCUS_STOP_COLOR

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���O�p�X�L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLog_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���O�p�X�t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtLog_LostFocus()

    imtLog.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�p�X���[�h�t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtPWD_GotFocus()

    imtPWD.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�p�X���[�h�L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtPWD_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�p�X���[�h�t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtPWD_LostFocus()

    imtPWD.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�T�[�o�[���t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtServer_GotFocus()

    imtServer.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�T�[�o�[���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtServer_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�T�[�o�[���t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtServer_LostFocus()

    imtServer.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�f�[�^�x�[�X���t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtSQLDBName_GotFocus()

    imtSQLDBName.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�f�[�^�x�[�X���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtSQLDBName_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�f�[�^�x�[�X���t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtSQLDBName_LostFocus()

    imtSQLDBName.BackColor = FOCUS_NO_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���[�U�[���t�H�[�J�X�擾��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtUID_GotFocus()

    imtUID.BackColor = FOCUS_STOP_COLOR
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���[�U�[���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtUID_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���[�U�[���t�H�[�J�X�r����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�T�^�Q�S
'�X�V�����@�F
'
Private Sub imtUID_LostFocus()

    imtUID.BackColor = FOCUS_NO_COLOR
    
End Sub
