VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmYpmf280 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   3885
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12480
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf280.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   12480
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame fraUketuke 
      Height          =   2175
      Left            =   240
      TabIndex        =   38
      Top             =   540
      Width           =   12015
      Begin VB.OptionButton optSyuukei 
         Caption         =   "�W�v���Ȃ�(����)"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3840
         TabIndex        =   8
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optSyuukei 
         Caption         =   "���i���ƂɏW�v"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Value           =   -1  'True
         Width           =   2115
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�Ώ۔N����"
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
      Begin imText6Ctl.imText txtUYear 
         Height          =   420
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Top             =   240
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":0CFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":0D68
         Key             =   "frmYpmf280.frx":0D86
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
      Begin imText6Ctl.imText txtUMonth 
         Height          =   420
         Index           =   0
         Left            =   3060
         TabIndex        =   2
         Top             =   240
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":0DBA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":0E28
         Key             =   "frmYpmf280.frx":0E46
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
      Begin imText6Ctl.imText txtUDay 
         Height          =   420
         Index           =   0
         Left            =   4260
         TabIndex        =   3
         Top             =   240
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":0E7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":0EE8
         Key             =   "frmYpmf280.frx":0F06
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
      Begin imText6Ctl.imText txtUYear 
         Height          =   420
         Index           =   1
         Left            =   6060
         TabIndex        =   4
         Top             =   240
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":0F3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":0FA8
         Key             =   "frmYpmf280.frx":0FC6
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
      Begin imText6Ctl.imText txtUMonth 
         Height          =   420
         Index           =   1
         Left            =   7500
         TabIndex        =   5
         Top             =   240
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":0FFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":1068
         Key             =   "frmYpmf280.frx":1086
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
      Begin imText6Ctl.imText txtUDay 
         Height          =   420
         Index           =   1
         Left            =   8700
         TabIndex        =   6
         Top             =   240
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":10BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":1128
         Key             =   "frmYpmf280.frx":1146
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
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   780
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�W�v����"
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
         Index           =   11
         Left            =   3720
         TabIndex        =   46
         Top             =   300
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
         Index           =   10
         Left            =   2580
         TabIndex        =   45
         Top             =   300
         Width           =   375
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
         Index           =   9
         Left            =   4860
         TabIndex        =   44
         Top             =   300
         Width           =   435
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
         Index           =   8
         Left            =   8160
         TabIndex        =   43
         Top             =   300
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
         Index           =   7
         Left            =   7020
         TabIndex        =   42
         Top             =   300
         Width           =   375
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
         Index           =   6
         Left            =   9300
         TabIndex        =   41
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   2  '��������
         Caption         =   "�`"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5340
         TabIndex        =   40
         Top             =   300
         Width           =   555
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   4260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   60
      TabIndex        =   36
      Top             =   4320
      Width           =   5175
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   3060
      Width           =   12255
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   17
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
         Picture         =   "frmYpmf280.frx":117A
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10440
         TabIndex        =   19
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
         Picture         =   "frmYpmf280.frx":1196
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   8700
         TabIndex        =   18
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "CSV�o��(F12)"
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
         rPic.left       =   0
         rPic.top        =   8
         rPic.right      =   0
         rPic.bottom     =   0
         rText.left      =   3
         rText.top       =   8
         rText.right     =   109
         rText.bottom    =   27
         Picture         =   "frmYpmf280.frx":12F0
      End
   End
   Begin VB.Frame fraKyoubai 
      Height          =   2115
      Left            =   240
      TabIndex        =   22
      Top             =   600
      Width           =   12015
      Begin CSComboLib.CSComboBox cboBcode 
         Height          =   405
         Index           =   0
         Left            =   1620
         TabIndex        =   9
         Top             =   180
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf280.frx":130C
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   180
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
      Begin CSComboLib.CSComboBox cboBcode 
         Height          =   405
         Index           =   1
         Left            =   1620
         TabIndex        =   10
         Top             =   1020
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf280.frx":1325
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�Ώ۔N����"
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
      Begin imText6Ctl.imText txtYear 
         Height          =   420
         Index           =   0
         Left            =   1620
         TabIndex        =   11
         Top             =   1560
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":133E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":13AC
         Key             =   "frmYpmf280.frx":13CA
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
         Index           =   0
         Left            =   3060
         TabIndex        =   12
         Top             =   1560
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":13FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":146C
         Key             =   "frmYpmf280.frx":148A
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
      Begin imText6Ctl.imText txtDay 
         Height          =   420
         Index           =   0
         Left            =   4260
         TabIndex        =   13
         Top             =   1560
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":14BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":152C
         Key             =   "frmYpmf280.frx":154A
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
      Begin imText6Ctl.imText txtYear 
         Height          =   420
         Index           =   1
         Left            =   6060
         TabIndex        =   14
         Top             =   1560
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":157E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":15EC
         Key             =   "frmYpmf280.frx":160A
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
         Index           =   1
         Left            =   7500
         TabIndex        =   15
         Top             =   1560
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":163E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":16AC
         Key             =   "frmYpmf280.frx":16CA
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
      Begin imText6Ctl.imText txtDay 
         Height          =   420
         Index           =   1
         Left            =   8700
         TabIndex        =   16
         Top             =   1560
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf280.frx":16FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf280.frx":176C
         Key             =   "frmYpmf280.frx":178A
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
      Begin VB.Label Label3 
         Alignment       =   2  '��������
         Caption         =   "�`"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5340
         TabIndex        =   35
         Top             =   1620
         Width           =   555
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
         Index           =   5
         Left            =   9300
         TabIndex        =   34
         Top             =   1620
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
         Index           =   4
         Left            =   7020
         TabIndex        =   33
         Top             =   1620
         Width           =   375
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
         Index           =   3
         Left            =   8160
         TabIndex        =   32
         Top             =   1620
         Width           =   435
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
         Index           =   2
         Left            =   4860
         TabIndex        =   31
         Top             =   1620
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
         Left            =   2580
         TabIndex        =   30
         Top             =   1620
         Width           =   375
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
         Left            =   3720
         TabIndex        =   29
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��������
         Caption         =   "�`"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   27
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblBcode_Name 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
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
         Height          =   435
         Index           =   1
         Left            =   2700
         TabIndex        =   26
         Top             =   1020
         Width           =   9195
      End
      Begin VB.Label lblBcode_Name 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
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
         Height          =   435
         Index           =   0
         Left            =   2700
         TabIndex        =   25
         Top             =   180
         Width           =   9195
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   12960
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf280.frx":17BE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf280.frx":182C
      Key             =   "frmYpmf280.frx":184A
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
      Left            =   12960
      TabIndex        =   20
      Top             =   180
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf280.frx":188E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf280.frx":18FC
      Key             =   "frmYpmf280.frx":191A
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
      TabIndex        =   21
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf280.frx":195E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf280.frx":19CC
      Key             =   "frmYpmf280.frx":19EA
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2835
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   5001
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��t�f�[�^"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�����f�[�^(�����܂�)"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmYpmf280"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typCsv_Record
    Field01 As String
    Field02 As String
    Field03 As String
    Field04 As String
    Field05 As String
    Field06 As String
    Field07 As String
    Field08 As String
    Field09 As String
    Field10 As String
    Field11 As String
    Field12 As String
    Field13 As String
End Type

Private Sub cboBcode_Click(Index As Integer)

    Call cboBcode_Validate(Index, False)
    
End Sub

Private Sub cboBcode_DropDown(Index As Integer)

    Call MakecboBcode(cboBcode(Index))
    
End Sub

Private Sub cboBcode_GotFocus(Index As Integer)

    cboBcode(Index).BackColor = FOCUS_STOP_COLOR
    cboBcode(Index).Tag = cboBcode(Index).Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboBcode_LostFocus(Index As Integer)
   
    cboBcode(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboBcode_Validate(Index As Integer, Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cboBcode_Validate_Err
    
    If Trim(cboBcode(Index).Text) = "" Then Exit Sub
    If cboBcode(Index).Tag = cboBcode(Index).Text Then Exit Sub
    
    lblBcode_Name(Index).Caption = ""
    
    With adoRecordset1
        '���Ӑ�}�X�^
        strSQL = "{call sp_MT070;2(" & Trim(cboBcode(Index).Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Fdiv")) Then
                If .Fields("Fdiv") = BUSINESS_DIV_BUYER Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    lblBcode_Name(Index).Caption = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If Index = 0 Then
        cboBcode(1).Text = cboBcode(0).Text
        lblBcode_Name(1).Caption = lblBcode_Name(0).Caption
    End If
    
    Exit Sub

cboBcode_Validate_Err:

    Call MsgBox("�t�H�[�J�X�ړ��O�G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBcode_Validate_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F��ʃN���A�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
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
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err
    
    If TabStrip1.Tabs(1).Selected = True Then
        '���̓`�F�b�N
        If DoValidationChecks1() = False Then Exit Sub
        'CSV�t�@�C�����擾
        If Get_SaveCsvFileName() = False Then Exit Sub
        'CSV�t�@�C���쐬
        If optSyuukei(0).Value = True Then
            If MakeCsvData1(Trim(txtFileName.Text)) = False Then Exit Sub
        ElseIf optSyuukei(1).Value = True Then
            If MakeCsvData1_2(Trim(txtFileName.Text)) = False Then Exit Sub
        End If
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        '���̓`�F�b�N
        If DoValidationChecks2() = False Then Exit Sub
        'CSV�t�@�C�����擾
        If Get_SaveCsvFileName() = False Then Exit Sub
        'CSV�t�@�C���쐬
        If MakeCsvData2(Trim(txtFileName.Text)) = False Then Exit Sub
    End If
    
    Call MsgBox("�I�����܂����B", vbInformation + vbOKOnly, "")

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
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
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
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
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
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "�f�[�^�b�r�u�o��"
    
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
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set g_clsAdoSQL = Nothing
    Set g_clsAdoAccess = Nothing
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
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        cboBcode(0).Text = ""
        cboBcode(0).Tag = ""
        cboBcode(1).Text = ""
        cboBcode(1).Tag = ""
        lblBcode_Name(0).Caption = ""
        lblBcode_Name(1).Caption = ""
        txtYear(0).Text = ""
        txtMonth(0).Text = ""
        txtDay(0).Text = ""
        txtYear(1).Text = ""
        txtMonth(1).Text = ""
        txtDay(1).Text = ""
        txtUYear(0).Text = ""
        txtUMonth(0).Text = ""
        txtUDay(0).Text = ""
        txtUYear(1).Text = ""
        txtUMonth(1).Text = ""
        txtUDay(1).Text = ""
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
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F
'
Private Function DoValidationChecks2() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks2_Err

    If Trim(cboBcode(0).Text) = "" Then
        cboBcode(0).SetFocus
        strErrMsg = "����R�[�h����͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(cboBcode(1).Text) = "" Then
        cboBcode(1).SetFocus
        strErrMsg = "����R�[�h����͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtYear(0).Text) = "" Then
        txtYear(0).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtMonth(0).Text) = "" Then
        txtMonth(0).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtDay(0).Text) = "" Then
        txtDay(0).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtYear(1).Text) = "" Then
        txtYear(1).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtMonth(1).Text) = "" Then
        txtMonth(1).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtDay(1).Text) = "" Then
        txtDay(1).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks2 = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks2 = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "���̓`�F�b�N")
    
    Exit Function
    
DoValidationChecks2_Err:

    DoValidationChecks2 = False
    Call MsgBox("���̓`�F�b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks2_Err")

End Function

'�ځ@�I�@�@�FCSV�t�@�C���쐬(����&����)
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F�Q�O�O�S�^�O�U�^�O�V�@���������܂߂�
'�@�@�@�@�@�@�Q�O�O�S�^�O�W�^�R�P�@�������͎d���P���Ǝd�����z��\��
'
Private Function MakeCsvData2(strFileName As String) As Boolean

    Dim strSQL As String
    Dim adoDT021 As New ADODB.Recordset
    Dim adoDT031 As New ADODB.Recordset
    Dim adoDT010 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim intFreefile1 As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim Csv_Rec As typCsv_Record

    On Error GoTo MakeCsvData2_Err
    
    MakeCsvData2 = False
    
    Screen.MousePointer = vbHourglass
    
    '�Ώ۔N����
    strDateFrom = txtYear(0).Text & Format(txtMonth(0).Text, "00") & Format(txtDay(0).Text, "00")
    strDateTo = txtYear(1).Text & Format(txtMonth(1).Text, "00") & Format(txtDay(1).Text, "00")
    
    '�t�@�C���쐬
    intFreefile1 = FreeFile
    Open strFileName For Output As intFreefile1
    
    '�^�C�g���쐬
    Write #intFreefile1, "����R�[�h", "���喼��", "�J�Ó�", _
                         "���i�R�[�h", "���i��", _
                         "����", "���z", "��t�ԍ�", "�o�i�Җ�", "", _
                         "�d���P��", "�d�����z"
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF280"
    g_clsAdoAccess.Connection.Execute strSQL

    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF280"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '�������׃f�[�^�I�[�v��(�ݐ�)
    strSQL = "{call sp_YPMF2801;2('" & strDateFrom & "','" & strDateTo & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT021.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT021.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT021.EOF
'        Csv_Rec.Field01 = IIf(IsNull(adoDT021.Fields("Bcode")), "", adoDT021.Fields("Bcode"))
'        Csv_Rec.Field03 = Global_Get_StringDay(left$(adoDT021.Fields("Ocode"), 8))
'        Csv_Rec.Field02 = Global_Get_Bname(g_clsAdoSQL, Csv_Rec.Field01, Csv_Rec.Field03, "")
'        Csv_Rec.Field04 = IIf(IsNull(adoDT021.Fields("Icode")), "", adoDT021.Fields("Icode"))
'        Csv_Rec.Field05 = IIf(IsNull(adoDT021.Fields("Iname")), "", adoDT021.Fields("Iname"))
'        Csv_Rec.Field06 = IIf(IsNull(adoDT021.Fields("Qty")), 0, adoDT021.Fields("Qty"))
'        Csv_Rec.Field07 = IIf(IsNull(adoDT021.Fields("Price")), 0, adoDT021.Fields("Price"))
'        Csv_Rec.Field08 = IIf(IsNull(adoDT021.Fields("Pnum")), "", adoDT021.Fields("Pnum"))
'
'        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02, Csv_Rec.Field03, Csv_Rec.Field04, _
'                             Csv_Rec.Field05, Csv_Rec.Field06, Csv_Rec.Field07, Csv_Rec.Field08
                
        wkRecordset.AddNew
        wkRecordset.Fields("����R�[�h") = IIf(IsNull(adoDT021.Fields("Bcode")), "", adoDT021.Fields("Bcode"))
        wkRecordset.Fields("�J�Ó�") = Global_Get_StringDay(left$(adoDT021.Fields("Ocode"), 8))
        wkRecordset.Fields("���喼��") = Global_Get_Bname(g_clsAdoSQL, wkRecordset.Fields("����R�[�h"), wkRecordset.Fields("�J�Ó�"), "")
        wkRecordset.Fields("���i�R�[�h") = IIf(IsNull(adoDT021.Fields("Icode")), "", adoDT021.Fields("Icode"))
        wkRecordset.Fields("���i��") = IIf(IsNull(adoDT021.Fields("Iname")), "", adoDT021.Fields("Iname"))
        wkRecordset.Fields("����") = IIf(IsNull(adoDT021.Fields("Qty")), 0, adoDT021.Fields("Qty"))
        wkRecordset.Fields("���z") = IIf(IsNull(adoDT021.Fields("Price")), 0, adoDT021.Fields("Price"))
        wkRecordset.Fields("��t�ԍ�") = IIf(IsNull(adoDT021.Fields("Pnum")), "", adoDT021.Fields("Pnum"))
                
        '��t�f�[�^(�ݐ�)
        If IsNumeric(wkRecordset.Fields("��t�ԍ�")) = True Then
            strSQL = "{call sp_YPMF2807;2('" & wkRecordset.Fields("�J�Ó�") & "'," & wkRecordset.Fields("��t�ԍ�") & ")}"
            adoDT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT010.EOF = False Then
                wkRecordset.Fields("�o�i�Җ�") = IIf(IsNull(adoDT010.Fields("Sname")), "", adoDT010.Fields("Sname"))
            Else
                wkRecordset.Fields("�o�i�Җ�") = ""
            End If
            adoDT010.Close
        Else
            wkRecordset.Fields("�o�i�Җ�") = ""
        End If
        
        wkRecordset.Fields("�敪") = ""
        wkRecordset.Fields("�d���P��") = Null
        wkRecordset.Fields("�d�����z") = Null
        wkRecordset.Update
                
        adoDT021.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData2_Cancel:
    Loop
    adoDT021.Close
    
    '�������׃f�[�^�I�[�v��
    strSQL = "{call sp_YPMF2801;1('" & strDateFrom & "','" & strDateTo & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT021.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT021.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT021.EOF
'        Csv_Rec.Field01 = IIf(IsNull(adoDT021.Fields("Bcode")), "", adoDT021.Fields("Bcode"))
'        Csv_Rec.Field03 = Global_Get_StringDay(left$(adoDT021.Fields("Ocode"), 8))
'        Csv_Rec.Field02 = Global_Get_Bname(g_clsAdoSQL, Csv_Rec.Field01, Csv_Rec.Field03, "")
'        Csv_Rec.Field04 = IIf(IsNull(adoDT021.Fields("Icode")), "", adoDT021.Fields("Icode"))
'        Csv_Rec.Field05 = IIf(IsNull(adoDT021.Fields("Iname")), "", adoDT021.Fields("Iname"))
'        Csv_Rec.Field06 = IIf(IsNull(adoDT021.Fields("Qty")), 0, adoDT021.Fields("Qty"))
'        Csv_Rec.Field07 = IIf(IsNull(adoDT021.Fields("Price")), 0, adoDT021.Fields("Price"))
'        Csv_Rec.Field08 = IIf(IsNull(adoDT021.Fields("Pnum")), "", adoDT021.Fields("Pnum"))
'
'        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02, Csv_Rec.Field03, Csv_Rec.Field04, _
'                             Csv_Rec.Field05, Csv_Rec.Field06, Csv_Rec.Field07, Csv_Rec.Field08
                
                
        wkRecordset.AddNew
        wkRecordset.Fields("����R�[�h") = IIf(IsNull(adoDT021.Fields("Bcode")), "", adoDT021.Fields("Bcode"))
        wkRecordset.Fields("�J�Ó�") = Global_Get_StringDay(left$(adoDT021.Fields("Ocode"), 8))
        wkRecordset.Fields("���喼��") = Global_Get_Bname(g_clsAdoSQL, wkRecordset.Fields("����R�[�h"), wkRecordset.Fields("�J�Ó�"), "")
        wkRecordset.Fields("���i�R�[�h") = IIf(IsNull(adoDT021.Fields("Icode")), "", adoDT021.Fields("Icode"))
        wkRecordset.Fields("���i��") = IIf(IsNull(adoDT021.Fields("Iname")), "", adoDT021.Fields("Iname"))
        wkRecordset.Fields("����") = IIf(IsNull(adoDT021.Fields("Qty")), 0, adoDT021.Fields("Qty"))
        wkRecordset.Fields("���z") = IIf(IsNull(adoDT021.Fields("Price")), 0, adoDT021.Fields("Price"))
        wkRecordset.Fields("��t�ԍ�") = IIf(IsNull(adoDT021.Fields("Pnum")), "", adoDT021.Fields("Pnum"))
        
        '��t�f�[�^
        If IsNumeric(wkRecordset.Fields("��t�ԍ�")) = True Then
            strSQL = "{call sp_YPMF2807;1('" & wkRecordset.Fields("�J�Ó�") & "'," & wkRecordset.Fields("��t�ԍ�") & ")}"
            adoDT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT010.EOF = False Then
                wkRecordset.Fields("�o�i�Җ�") = IIf(IsNull(adoDT010.Fields("Sname")), "", adoDT010.Fields("Sname"))
            Else
                wkRecordset.Fields("�o�i�Җ�") = ""
            End If
            adoDT010.Close
        Else
            wkRecordset.Fields("�o�i�Җ�") = ""
        End If
        
        wkRecordset.Fields("�敪") = ""
        wkRecordset.Fields("�d���P��") = Null
        wkRecordset.Fields("�d�����z") = Null
        wkRecordset.Update
                
        adoDT021.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData2_Cancel:
    Loop
    adoDT021.Close
    
    '�������׃f�[�^�I�[�v��(�ݐ�)
    strSQL = "{call sp_YPMF2806;2('" & Format$(strDateFrom, "####/##/##") & "','" & Format$(strDateTo, "####/##/##") & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    adoDT031.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT031.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT031.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT031.EOF
        wkRecordset.AddNew
        wkRecordset.Fields("����R�[�h") = IIf(IsNull(adoDT031.Fields("Bcode")), "", adoDT031.Fields("Bcode"))
        wkRecordset.Fields("�J�Ó�") = adoDT031.Fields("Odate")
        wkRecordset.Fields("���喼��") = Global_Get_Bname(g_clsAdoSQL, wkRecordset.Fields("����R�[�h"), wkRecordset.Fields("�J�Ó�"), "")
        wkRecordset.Fields("���i�R�[�h") = IIf(IsNull(adoDT031.Fields("Icode")), "", adoDT031.Fields("Icode"))
        wkRecordset.Fields("���i��") = IIf(IsNull(adoDT031.Fields("Iname")), "", adoDT031.Fields("Iname"))
        wkRecordset.Fields("����") = IIf(IsNull(adoDT031.Fields("Qty")), 0, adoDT031.Fields("Qty"))
        wkRecordset.Fields("���z") = IIf(IsNull(adoDT031.Fields("Price")), 0, adoDT031.Fields("Price"))
        wkRecordset.Fields("��t�ԍ�") = adoDT031.Fields("Onum")
        wkRecordset.Fields("�o�i�Җ�") = IIf(IsNull(adoDT031.Fields("Sname")), "", adoDT031.Fields("Sname"))
        wkRecordset.Fields("�敪") = "*"
        wkRecordset.Fields("�d���P��") = IIf(IsNull(adoDT031.Fields("Price2")), 0, adoDT031.Fields("Price2"))
        wkRecordset.Fields("�d�����z") = Fix(CDec(wkRecordset.Fields("����")) * CDec(wkRecordset.Fields("�d���P��")))
        wkRecordset.Update
                
        adoDT031.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData2_Cancel:
    Loop
    adoDT031.Close
    
    '�������׃f�[�^�I�[�v��
    strSQL = "{call sp_YPMF2806;1('" & Format$(strDateFrom, "####/##/##") & "','" & Format$(strDateTo, "####/##/##") & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    adoDT031.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT031.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT031.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT031.EOF
        wkRecordset.AddNew
        wkRecordset.Fields("����R�[�h") = IIf(IsNull(adoDT031.Fields("Bcode")), "", adoDT031.Fields("Bcode"))
        wkRecordset.Fields("�J�Ó�") = adoDT031.Fields("Odate")
        wkRecordset.Fields("���喼��") = Global_Get_Bname(g_clsAdoSQL, wkRecordset.Fields("����R�[�h"), wkRecordset.Fields("�J�Ó�"), "")
        wkRecordset.Fields("���i�R�[�h") = IIf(IsNull(adoDT031.Fields("Icode")), "", adoDT031.Fields("Icode"))
        wkRecordset.Fields("���i��") = IIf(IsNull(adoDT031.Fields("Iname")), "", adoDT031.Fields("Iname"))
        wkRecordset.Fields("����") = IIf(IsNull(adoDT031.Fields("Qty")), 0, adoDT031.Fields("Qty"))
        wkRecordset.Fields("���z") = IIf(IsNull(adoDT031.Fields("Price")), 0, adoDT031.Fields("Price"))
        wkRecordset.Fields("��t�ԍ�") = adoDT031.Fields("Onum")
        wkRecordset.Fields("�o�i�Җ�") = IIf(IsNull(adoDT031.Fields("Sname")), "", adoDT031.Fields("Sname"))
        wkRecordset.Fields("�敪") = "*"
        wkRecordset.Fields("�d���P��") = IIf(IsNull(adoDT031.Fields("Price2")), 0, adoDT031.Fields("Price2"))
        wkRecordset.Fields("�d�����z") = Fix(CDec(wkRecordset.Fields("����")) * CDec(wkRecordset.Fields("�d���P��")))
        wkRecordset.Update
                
        adoDT031.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData2_Cancel:
    Loop
    adoDT031.Close
    
    wkRecordset.Close
    
    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF280 ORDER BY ����R�[�h,�J�Ó�,�敪,��t�ԍ�"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    Do While Not wkRecordset.EOF
        Csv_Rec.Field01 = IIf(IsNull(wkRecordset.Fields("����R�[�h")), "", wkRecordset.Fields("����R�[�h"))
        Csv_Rec.Field02 = IIf(IsNull(wkRecordset.Fields("���喼��")), "", wkRecordset.Fields("���喼��"))
        Csv_Rec.Field03 = IIf(IsNull(wkRecordset.Fields("�J�Ó�")), "", wkRecordset.Fields("�J�Ó�"))
        Csv_Rec.Field04 = IIf(IsNull(wkRecordset.Fields("���i�R�[�h")), "", wkRecordset.Fields("���i�R�[�h"))
        Csv_Rec.Field05 = IIf(IsNull(wkRecordset.Fields("���i��")), "", wkRecordset.Fields("���i��"))
        Csv_Rec.Field06 = IIf(IsNull(wkRecordset.Fields("����")), 0, wkRecordset.Fields("����"))
        Csv_Rec.Field07 = IIf(IsNull(wkRecordset.Fields("���z")), 0, wkRecordset.Fields("���z"))
        Csv_Rec.Field08 = IIf(IsNull(wkRecordset.Fields("��t�ԍ�")), "", wkRecordset.Fields("��t�ԍ�"))
        Csv_Rec.Field09 = IIf(IsNull(wkRecordset.Fields("�o�i�Җ�")), "", wkRecordset.Fields("�o�i�Җ�"))
        Csv_Rec.Field10 = IIf(IsNull(wkRecordset.Fields("�敪")), "", wkRecordset.Fields("�敪"))
        
        Csv_Rec.Field11 = IIf(IsNull(wkRecordset.Fields("�d���P��")), "", wkRecordset.Fields("�d���P��"))
        Csv_Rec.Field12 = IIf(IsNull(wkRecordset.Fields("�d�����z")), "", wkRecordset.Fields("�d�����z"))

        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02, Csv_Rec.Field03, Csv_Rec.Field04, Csv_Rec.Field05, _
                             Csv_Rec.Field06, Csv_Rec.Field07, Csv_Rec.Field08, Csv_Rec.Field09, Csv_Rec.Field10, _
                             Csv_Rec.Field11, Csv_Rec.Field12

        wkRecordset.MoveNext
    Loop
    wkRecordset.Close
    
    MakeCsvData2 = True
    
MakeCsvData2_Exit:
    
    Close
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    DoEvents
    
    Exit Function

MakeCsvData2_Cancel:

    GoTo MakeCsvData2_Exit:

MakeCsvData2_Err:

    MakeCsvData2 = False
    Call MsgBox("CSV�t�@�C���쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeCsvData2_Err")
    GoTo MakeCsvData2_Exit:

End Function

'�ځ@�I�@�@�F�R���{�{�b�N�X�̍쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
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
'                If .Fields("Sdate") <= Trim(lblOdate.Caption) And Trim(lblOdate.Caption) <= .Fields("Fdate") Then
'                    Ctrl.AddItem .Fields("Bnum") & ";" & .Fields("Bname")
'                Else
                    Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
'                End If
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

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    If TabStrip1.Tabs(1).Selected = True Then
        txtUYear(0).SetFocus
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        cboBcode(0).SetFocus
    End If

End Sub

Private Sub TabStrip1_Click()

    On Error Resume Next

    If TabStrip1.Tabs(1).Selected = True Then
        fraUketuke.Visible = True
        fraKyoubai.Visible = False
        txtUYear(0).SetFocus
        DoEvents
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        fraUketuke.Visible = False
        fraKyoubai.Visible = True
        cboBcode(0).SetFocus
        DoEvents
    End If

End Sub

Private Sub txtDay_GotFocus(Index As Integer)
    
    txtDay(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtDay_LostFocus(Index As Integer)
    
    txtDay(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtMonth_GotFocus(Index As Integer)
    
    txtMonth(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtMonth_LostFocus(Index As Integer)
    
    txtMonth(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtUDay_GotFocus(Index As Integer)
    
    txtUDay(Index).BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtUDay_LostFocus(Index As Integer)
    
    txtUDay(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtUMonth_GotFocus(Index As Integer)
   
    txtUMonth(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtUMonth_LostFocus(Index As Integer)
    
    txtUMonth(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtUYear_GotFocus(Index As Integer)
    
    txtUYear(Index).BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtUYear_LostFocus(Index As Integer)
    
    txtUYear(Index).BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtYear_GotFocus(Index As Integer)
    
    txtYear(Index).BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtYear_LostFocus(Index As Integer)
    
    txtYear(Index).BackColor = FOCUS_NO_COLOR

End Sub

Private Function Get_SaveCsvFileName() As Boolean

    Dim strSQL As String

    On Error GoTo Get_SaveCsvFileName_Err
    
    Get_SaveCsvFileName = False
    txtFileName.Text = ""
    
    With CommonDialog1
        .DialogTitle = "csv̧�ق��w��"
        .FileName = ""
        .CancelError = False
        .Filter = "csv̧�� (*.csv)|*.csv|���ׂĂ�̧�� (*.*)|*.*|"
        '.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer
        '�R�����_�C�A���O�{�b�N�X���J��
        .ShowSave
        If Len(.FileName) = 0 Then Exit Function
        '�t�@�C�����擾
        txtFileName.Text = .FileName
    End With
    
    If Trim(txtFileName.Text) = "" Then Exit Function
    
    '���Ƀt�@�C��������ꍇ
    If Dir(txtFileName.Text) <> "" Then
        If MsgBox("�㏑�����܂����H", vbInformation + vbYesNo, "") = vbNo Then Exit Function
    End If
    
    Get_SaveCsvFileName = True
    
    Exit Function
    
Get_SaveCsvFileName_Err:

    Get_SaveCsvFileName = False
    Call MsgBox("CSV�t�@�C�����擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Get_SaveCsvFileName_Err")

End Function

'�ځ@�I�@�@�F���̓`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F
'
Private Function DoValidationChecks1() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks1_Err

    If Trim(txtUYear(0).Text) = "" Then
        txtUYear(0).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtUMonth(0).Text) = "" Then
        txtUMonth(0).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtUDay(0).Text) = "" Then
        txtUDay(0).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtUYear(1).Text) = "" Then
        txtUYear(1).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtUMonth(1).Text) = "" Then
        txtUMonth(1).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(txtUDay(1).Text) = "" Then
        txtUDay(1).SetFocus
        strErrMsg = "�Ώ۔N��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks1 = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks1 = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "���̓`�F�b�N")
    
    Exit Function
    
DoValidationChecks1_Err:

    DoValidationChecks1 = False
    Call MsgBox("���̓`�F�b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks1_Err")

End Function

'�ځ@�I�@�@�FCSV�t�@�C���쐬(��t)
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F
'
Private Function MakeCsvData1(strFileName As String) As Boolean

    Dim strSQL As String
    Dim adoDT011 As New ADODB.Recordset
    Dim intFreefile1 As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim Csv_Rec As typCsv_Record

    On Error GoTo MakeCsvData1_Err
    
    MakeCsvData1 = False
    
    Screen.MousePointer = vbHourglass
    
    '�Ώ۔N����
    strDateFrom = txtUYear(0).Text & "/" & Format(txtUMonth(0).Text, "00") & "/" & Format(txtUDay(0).Text, "00")
    strDateTo = txtUYear(1).Text & "/" & Format(txtUMonth(1).Text, "00") & "/" & Format(txtUDay(1).Text, "00")
    
    '�t�@�C���쐬
    intFreefile1 = FreeFile
    Open strFileName For Output As intFreefile1
    
    '�^�C�g���쐬
    Write #intFreefile1, "���i��", "����"
    
    '��t���׃f�[�^�I�[�v��(�ݐ�)
    If optSyuukei(0).Value = True Then
        strSQL = "{call sp_YPMF2802;2('" & strDateFrom & "','" & strDateTo & "'" & ")}"
    ElseIf optSyuukei(1).Value = True Then
        strSQL = "{call sp_YPMF2803;2('" & strDateFrom & "','" & strDateTo & "'" & ")}"
    End If
    adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT011.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT011.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT011.EOF
        Csv_Rec.Field01 = IIf(IsNull(adoDT011.Fields("Iname")), "", adoDT011.Fields("Iname"))
        Csv_Rec.Field02 = IIf(IsNull(adoDT011.Fields("Qty")), "", adoDT011.Fields("Qty"))
    
        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02
                
        adoDT011.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData1_Cancel:
    Loop
    adoDT011.Close
    
    '�������׃f�[�^�I�[�v��
    If optSyuukei(0).Value = True Then
        strSQL = "{call sp_YPMF2802;1('" & strDateFrom & "','" & strDateTo & "'" & ")}"
    ElseIf optSyuukei(1).Value = True Then
        strSQL = "{call sp_YPMF2803;1('" & strDateFrom & "','" & strDateTo & "'" & ")}"
    End If
    adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT011.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT011.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT011.EOF
        Csv_Rec.Field01 = IIf(IsNull(adoDT011.Fields("Iname")), "", Trim(adoDT011.Fields("Iname")))
        Csv_Rec.Field02 = IIf(IsNull(adoDT011.Fields("Qty")), "", adoDT011.Fields("Qty"))
    
        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02
                
        adoDT011.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData1_Cancel:
    Loop
    adoDT011.Close
    
    MakeCsvData1 = True
    
MakeCsvData1_Exit:
    
    Close
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    DoEvents
    
    Exit Function

MakeCsvData1_Cancel:

    GoTo MakeCsvData1_Exit:

MakeCsvData1_Err:

    MakeCsvData1 = False
    Call MsgBox("CSV�t�@�C���쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeCsvData1_Err")
    GoTo MakeCsvData1_Exit:

End Function

'�ځ@�I�@�@�FCSV�t�@�C���쐬(��t)
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�P�Q�^�P�U
'�X�V�����@�F
'
Private Function MakeCsvData1_2(strFileName As String) As Boolean

    Dim strSQL As String
    Dim adoDT011 As New ADODB.Recordset
    Dim adoDT021 As New ADODB.Recordset
    Dim intFreefile1 As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim Csv_Rec As typCsv_Record

    On Error GoTo MakeCsvData1_2_Err
    
    MakeCsvData1_2 = False
    
    Screen.MousePointer = vbHourglass
    
    '�Ώ۔N����
    strDateFrom = txtUYear(0).Text & "/" & Format(txtUMonth(0).Text, "00") & "/" & Format(txtUDay(0).Text, "00")
    strDateTo = txtUYear(1).Text & "/" & Format(txtUMonth(1).Text, "00") & "/" & Format(txtUDay(1).Text, "00")
    
    '�t�@�C���쐬
    intFreefile1 = FreeFile
    Open strFileName For Output As intFreefile1
    
    '�^�C�g���쐬
    '�^�C�g���쐬
    Write #intFreefile1, "�J�Ó�", "��t�ԍ�", "�o�i�Җ�", "�s", "���i�R�[�h", _
                         "���i��", "����", "���z", "����R�[�h", "���喼��"
    
    '��t���׃f�[�^�I�[�v��(�ݐ�)
    strSQL = "{call sp_YPMF2804;2('" & strDateFrom & "','" & strDateTo & "'" & ")}"
    adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT011.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT011.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT011.EOF
        Csv_Rec.Field01 = IIf(IsNull(adoDT011.Fields("Odate")), "", adoDT011.Fields("Odate"))
        Csv_Rec.Field02 = IIf(IsNull(adoDT011.Fields("Pnum")), "", adoDT011.Fields("Pnum"))
        Csv_Rec.Field03 = IIf(IsNull(adoDT011.Fields("Sname")), "", adoDT011.Fields("Sname"))
        Csv_Rec.Field04 = IIf(IsNull(adoDT011.Fields("Line")), "", adoDT011.Fields("Line"))
        Csv_Rec.Field05 = IIf(IsNull(adoDT011.Fields("Icode")), "", adoDT011.Fields("Icode"))
        Csv_Rec.Field06 = IIf(IsNull(adoDT011.Fields("Iname")), "", adoDT011.Fields("Iname"))
        Csv_Rec.Field07 = IIf(IsNull(adoDT011.Fields("Qty")), "0", adoDT011.Fields("Qty"))

        '�������׃f�[�^�I�[�v��(�ݐ�)
        strSQL = "{call sp_YPMF2805;2('" & Global_Get_NumericDay(adoDT011.Fields("Odate")) & "'," & _
                 adoDT011.Fields("Pnum") & "," & adoDT011.Fields("Line") & ")}"
        adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoDT021.EOF = False Then
            Csv_Rec.Field08 = IIf(IsNull(adoDT021.Fields("Price")), "0", adoDT021.Fields("Price"))
            Csv_Rec.Field09 = IIf(IsNull(adoDT021.Fields("Bcode")), "", adoDT021.Fields("Bcode"))
            Csv_Rec.Field10 = IIf(IsNull(adoDT021.Fields("Bname")), "", Trim(adoDT021.Fields("Bname")))
            If IsNull(adoDT021.Fields("Sline")) = False Then
                If adoDT021.Fields("Sline") <> 0 Then
                    Csv_Rec.Field11 = "���Z"
                End If
            End If
            If IsNull(adoDT021.Fields("Idiv")) = False Then
                If adoDT021.Fields("Idiv") = 1 Then
                    Csv_Rec.Field11 = "�s����"
                End If
            End If
        Else
            Csv_Rec.Field08 = "0"
            Csv_Rec.Field09 = ""
            Csv_Rec.Field10 = ""
            Csv_Rec.Field11 = ""
        End If
        adoDT021.Close
    
        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02, Csv_Rec.Field03, Csv_Rec.Field04, Csv_Rec.Field05, _
                             Csv_Rec.Field06, Csv_Rec.Field07, Csv_Rec.Field08, Csv_Rec.Field09, Csv_Rec.Field10
                
        adoDT011.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData1_2_Cancel:
    Loop
    adoDT011.Close
    
    '��t���׃f�[�^�I�[�v��
    strSQL = "{call sp_YPMF2804;1('" & strDateFrom & "','" & strDateTo & "'" & ")}"
    adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT011.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT011.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT011.EOF
        Csv_Rec.Field01 = IIf(IsNull(adoDT011.Fields("Odate")), "", adoDT011.Fields("Odate"))
        Csv_Rec.Field02 = IIf(IsNull(adoDT011.Fields("Pnum")), "", adoDT011.Fields("Pnum"))
        Csv_Rec.Field03 = IIf(IsNull(adoDT011.Fields("Sname")), "", adoDT011.Fields("Sname"))
        Csv_Rec.Field04 = IIf(IsNull(adoDT011.Fields("Line")), "", adoDT011.Fields("Line"))
        Csv_Rec.Field05 = IIf(IsNull(adoDT011.Fields("Icode")), "", adoDT011.Fields("Icode"))
        Csv_Rec.Field06 = IIf(IsNull(adoDT011.Fields("Iname")), "", adoDT011.Fields("Iname"))
        Csv_Rec.Field07 = IIf(IsNull(adoDT011.Fields("Qty")), "0", adoDT011.Fields("Qty"))

        '�������׃f�[�^�I�[�v��
        strSQL = "{call sp_YPMF2805;1('" & Global_Get_NumericDay(adoDT011.Fields("Odate")) & "'," & _
                 adoDT011.Fields("Pnum") & "," & adoDT011.Fields("Line") & ")}"
        adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoDT021.EOF = False Then
            Csv_Rec.Field08 = IIf(IsNull(adoDT021.Fields("Price")), "0", adoDT021.Fields("Price"))
            Csv_Rec.Field09 = IIf(IsNull(adoDT021.Fields("Bcode")), "", adoDT021.Fields("Bcode"))
            Csv_Rec.Field10 = IIf(IsNull(adoDT021.Fields("Bname")), "", Trim(adoDT021.Fields("Bname")))
            If IsNull(adoDT021.Fields("Sline")) = False Then
                If adoDT021.Fields("Sline") <> 0 Then
                    Csv_Rec.Field11 = "���Z"
                End If
            End If
            If IsNull(adoDT021.Fields("Idiv")) = False Then
                If adoDT021.Fields("Idiv") = 1 Then
                    Csv_Rec.Field11 = "�s����"
                End If
            End If
        Else
            Csv_Rec.Field08 = "0"
            Csv_Rec.Field09 = ""
            Csv_Rec.Field10 = ""
            Csv_Rec.Field11 = ""
        End If
        adoDT021.Close
    
        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02, Csv_Rec.Field03, Csv_Rec.Field04, Csv_Rec.Field05, _
                             Csv_Rec.Field06, Csv_Rec.Field07, Csv_Rec.Field08, Csv_Rec.Field09, Csv_Rec.Field10
                
        adoDT011.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData1_2_Cancel:
    Loop
    adoDT011.Close
    
    MakeCsvData1_2 = True
    
MakeCsvData1_2_Exit:
    
    Close
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    DoEvents
    
    Exit Function

MakeCsvData1_2_Cancel:

    GoTo MakeCsvData1_2_Exit:

MakeCsvData1_2_Err:

    MakeCsvData1_2 = False
    Call MsgBox("CSV�t�@�C���쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeCsvData1_2_Err")
    GoTo MakeCsvData1_2_Exit:

End Function

'�ځ@�I�@�@�FCSV�t�@�C���쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�Q�Q
'�X�V�����@�F
'
Private Function MakeCsvData2_Old(strFileName As String) As Boolean

    Dim strSQL As String
    Dim adoDT021 As New ADODB.Recordset
    Dim intFreefile1 As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim Csv_Rec As typCsv_Record

    On Error GoTo MakeCsvData2_Old_Err
    
    MakeCsvData2_Old = False
    
    Screen.MousePointer = vbHourglass
    
    '�Ώ۔N����
    strDateFrom = txtYear(0).Text & Format(txtMonth(0).Text, "00") & Format(txtDay(0).Text, "00")
    strDateTo = txtYear(1).Text & Format(txtMonth(1).Text, "00") & Format(txtDay(1).Text, "00")
    
    '�t�@�C���쐬
    intFreefile1 = FreeFile
    Open strFileName For Output As intFreefile1
    
    '�^�C�g���쐬
    Write #intFreefile1, "����R�[�h", "���喼��", "�J�Ó�", _
                         "���i�R�[�h", "���i��", _
                         "����", "���z", "��t�ԍ�"
    
    '�������׃f�[�^�I�[�v��(�ݐ�)
    strSQL = "{call sp_YPMF2801;2('" & strDateFrom & "','" & strDateTo & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT021.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT021.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT021.EOF
        Csv_Rec.Field01 = IIf(IsNull(adoDT021.Fields("Bcode")), "", adoDT021.Fields("Bcode"))
        Csv_Rec.Field03 = Global_Get_StringDay(left$(adoDT021.Fields("Ocode"), 8))
        Csv_Rec.Field02 = Global_Get_Bname(g_clsAdoSQL, Csv_Rec.Field01, Csv_Rec.Field03, "")
        Csv_Rec.Field04 = IIf(IsNull(adoDT021.Fields("Icode")), "", adoDT021.Fields("Icode"))
        Csv_Rec.Field05 = IIf(IsNull(adoDT021.Fields("Iname")), "", adoDT021.Fields("Iname"))
        Csv_Rec.Field06 = IIf(IsNull(adoDT021.Fields("Qty")), 0, adoDT021.Fields("Qty"))
        Csv_Rec.Field07 = IIf(IsNull(adoDT021.Fields("Price")), 0, adoDT021.Fields("Price"))
        Csv_Rec.Field08 = IIf(IsNull(adoDT021.Fields("Pnum")), "", adoDT021.Fields("Pnum"))
    
        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02, Csv_Rec.Field03, Csv_Rec.Field04, _
                             Csv_Rec.Field05, Csv_Rec.Field06, Csv_Rec.Field07, Csv_Rec.Field08
                
        adoDT021.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData2_Old_Cancel:
    Loop
    adoDT021.Close
    
    '�������׃f�[�^�I�[�v��
    strSQL = "{call sp_YPMF2801;1('" & strDateFrom & "','" & strDateTo & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT021.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT021.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    Do While Not adoDT021.EOF
        Csv_Rec.Field01 = IIf(IsNull(adoDT021.Fields("Bcode")), "", adoDT021.Fields("Bcode"))
        Csv_Rec.Field03 = Global_Get_StringDay(left$(adoDT021.Fields("Ocode"), 8))
        Csv_Rec.Field02 = Global_Get_Bname(g_clsAdoSQL, Csv_Rec.Field01, Csv_Rec.Field03, "")
        Csv_Rec.Field04 = IIf(IsNull(adoDT021.Fields("Icode")), "", adoDT021.Fields("Icode"))
        Csv_Rec.Field05 = IIf(IsNull(adoDT021.Fields("Iname")), "", adoDT021.Fields("Iname"))
        Csv_Rec.Field06 = IIf(IsNull(adoDT021.Fields("Qty")), 0, adoDT021.Fields("Qty"))
        Csv_Rec.Field07 = IIf(IsNull(adoDT021.Fields("Price")), 0, adoDT021.Fields("Price"))
        Csv_Rec.Field08 = IIf(IsNull(adoDT021.Fields("Pnum")), "", adoDT021.Fields("Pnum"))
    
        Write #intFreefile1, Csv_Rec.Field01, Csv_Rec.Field02, Csv_Rec.Field03, Csv_Rec.Field04, _
                             Csv_Rec.Field05, Csv_Rec.Field06, Csv_Rec.Field07, Csv_Rec.Field08
                
        adoDT021.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakeCsvData2_Old_Cancel:
    Loop
    adoDT021.Close
    
    MakeCsvData2_Old = True
    
MakeCsvData2_Old_Exit:
    
    Close
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    DoEvents
    
    Exit Function

MakeCsvData2_Old_Cancel:

    GoTo MakeCsvData2_Old_Exit:

MakeCsvData2_Old_Err:

    MakeCsvData2_Old = False
    Call MsgBox("CSV�t�@�C���쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeCsvData2_Old_Err")
    GoTo MakeCsvData2_Old_Exit:

End Function
