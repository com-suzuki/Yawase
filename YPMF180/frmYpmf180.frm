VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmYpmf180 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   3015
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
   Icon            =   "frmYpmf180.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   12150
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   17
      Top             =   2220
      Width           =   12015
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   13
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
         Picture         =   "frmYpmf180.frx":0CFA
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10200
         TabIndex        =   12
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
         Picture         =   "frmYpmf180.frx":0D16
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   8460
         TabIndex        =   11
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
         Picture         =   "frmYpmf180.frx":0E70
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   2235
      Left            =   60
      TabIndex        =   16
      Top             =   0
      Width           =   12015
      Begin VB.OptionButton Option1 
         Caption         =   "�A�ؖ��̏�"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   5040
         TabIndex        =   10
         Top             =   1620
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�o�i�ҁA�J�Ó�����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   1620
         TabIndex        =   9
         Top             =   1620
         Value           =   -1  'True
         Width           =   2955
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�o�i�Җ�"
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
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "���t�͈�"
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
      Begin imNumber6Ctl.imNumber imnRdate_Year 
         Height          =   375
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Top             =   240
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   661
         Calculator      =   "frmYpmf180.frx":0F82
         Caption         =   "frmYpmf180.frx":0FA2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf180.frx":1010
         Keys            =   "frmYpmf180.frx":102E
         Spin            =   "frmYpmf180.frx":1068
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2099
         MinValue        =   1900
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   2099
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnRdate_Month 
         Height          =   375
         Index           =   0
         Left            =   2700
         TabIndex        =   2
         Top             =   240
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   661
         Calculator      =   "frmYpmf180.frx":1090
         Caption         =   "frmYpmf180.frx":10B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf180.frx":111E
         Keys            =   "frmYpmf180.frx":113C
         Spin            =   "frmYpmf180.frx":1176
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   12
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnRdate_Day 
         Height          =   375
         Index           =   0
         Left            =   3540
         TabIndex        =   3
         Top             =   240
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmYpmf180.frx":119E
         Caption         =   "frmYpmf180.frx":11BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf180.frx":122C
         Keys            =   "frmYpmf180.frx":124A
         Spin            =   "frmYpmf180.frx":1284
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   31
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   31
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnRdate_Year 
         Height          =   375
         Index           =   1
         Left            =   4740
         TabIndex        =   4
         Top             =   240
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   661
         Calculator      =   "frmYpmf180.frx":12AC
         Caption         =   "frmYpmf180.frx":12CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf180.frx":133A
         Keys            =   "frmYpmf180.frx":1358
         Spin            =   "frmYpmf180.frx":1392
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2099
         MinValue        =   1900
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   2099
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnRdate_Month 
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   5
         Top             =   240
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   661
         Calculator      =   "frmYpmf180.frx":13BA
         Caption         =   "frmYpmf180.frx":13DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf180.frx":1448
         Keys            =   "frmYpmf180.frx":1466
         Spin            =   "frmYpmf180.frx":14A0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   12
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnRdate_Day 
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   6
         Top             =   240
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmYpmf180.frx":14C8
         Caption         =   "frmYpmf180.frx":14E8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf180.frx":1556
         Keys            =   "frmYpmf180.frx":1574
         Spin            =   "frmYpmf180.frx":15AE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   31
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   31
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
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
      Begin imText6Ctl.imText txtIname 
         Height          =   345
         Left            =   1620
         TabIndex        =   7
         Top             =   720
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   609
         Caption         =   "frmYpmf180.frx":15D6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf180.frx":1644
         Key             =   "frmYpmf180.frx":1662
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
         MaxLength       =   40
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWW"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   4
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText txtSname 
         Height          =   345
         Left            =   1620
         TabIndex        =   8
         Top             =   1200
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   609
         Caption         =   "frmYpmf180.frx":1696
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf180.frx":1704
         Key             =   "frmYpmf180.frx":1722
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
         MaxLength       =   40
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWW"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   4
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
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�o�͏�"
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
      Begin VB.Label Label1 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "(�o�i�Җ��̈ꕔ����͂��Ă�������)"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   5100
         TabIndex        =   29
         Top             =   1200
         Width           =   4935
      End
      Begin VB.Label Label1 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "(�A�ؖ��̈ꕔ����͂��Ă�������)"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   5100
         TabIndex        =   28
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�N"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   5580
         TabIndex        =   26
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   6420
         TabIndex        =   25
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   7140
         TabIndex        =   24
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�N"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   30
         Left            =   2400
         TabIndex        =   23
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   31
         Left            =   3180
         TabIndex        =   22
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   32
         Left            =   3960
         TabIndex        =   21
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "�`"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4380
         TabIndex        =   20
         Top             =   300
         Width           =   375
      End
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
      Caption         =   "frmYpmf180.frx":1756
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf180.frx":17C4
      Key             =   "frmYpmf180.frx":17E2
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
      TabIndex        =   14
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf180.frx":1826
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf180.frx":1894
      Key             =   "frmYpmf180.frx":18B2
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
      TabIndex        =   15
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf180.frx":18F6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf180.frx":1964
      Key             =   "frmYpmf180.frx":1982
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
Attribute VB_Name = "frmYpmf180"
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
'�X�V�����@�F
'
Private Sub cmdClear_Click()

    Call FieldsClear
    imnRdate_Year(0).SetFocus

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "�A�؈ꗗ�\"

    Call FieldsClear
    
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
'�X�V�����@�F
'
Private Sub FieldsClear()

    On Error GoTo FieldsClear_Err
    
    imnRdate_Year(0).Value = "2002"  '2002�N9��1���ɃV�X�e������
    imnRdate_Month(0).Value = "9"
    imnRdate_Day(0).Value = "1"
    imnRdate_Year(1).Value = Year(Now())
    imnRdate_Month(1).Value = Month(Now())
    imnRdate_Day(1).Value = Day(Now())
    txtIname.Text = ""
    txtSname.Text = ""
    
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
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
'�X�V�����@�F
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(imnRdate_Year(0).Text) = "" Then
        imnRdate_Year(0).SetFocus
        strErrMsg = "�N����͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Month(0).Text) = "" Then
        imnRdate_Month(0).SetFocus
        strErrMsg = "������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Day(0).Text) = "" Then
        imnRdate_Day(0).SetFocus
        strErrMsg = "������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Year(1).Text) = "" Then
        imnRdate_Year(1).SetFocus
        strErrMsg = "�N����͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Month(1).Text) = "" Then
        imnRdate_Month(1).SetFocus
        strErrMsg = "������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Day(1).Text) = "" Then
        imnRdate_Day(1).SetFocus
        strErrMsg = "������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    
    If Len(Trim(imnRdate_Year(0).Text)) < 4 Then
        imnRdate_Year(0).SetFocus
        strErrMsg = "�N�𐼗�œ��͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Len(Trim(imnRdate_Year(1).Text)) < 4 Then
        imnRdate_Year(1).SetFocus
        strErrMsg = "�N�𐼗�œ��͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    
    If IsDate(imnRdate_Year(0).Text & "/" & Format$(imnRdate_Month(0).Text, "00") & "/" & Format$(imnRdate_Day(0).Text, "00")) = False Then
        imnRdate_Day(0).SetFocus
        strErrMsg = "���������t����͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If IsDate(imnRdate_Year(1).Text & "/" & Format$(imnRdate_Month(1).Text, "00") & "/" & Format$(imnRdate_Day(1).Text, "00")) = False Then
        imnRdate_Day(1).SetFocus
        strErrMsg = "���������t����͂��Ă��������B"
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

'�ځ@�I�@�@�F����p���[�N�쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
'�X�V�����@�F
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim adoDT011 As New ADODB.Recordset
    Dim adoDT021 As New ADODB.Recordset
    Dim adoDT031 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim strFromDate As String
    Dim strToDate As String

    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    strFromDate = imnRdate_Year(0).Text & "/" & Format$(imnRdate_Month(0).Text, "00") & "/" & Format$(imnRdate_Day(0).Text, "00")
    strToDate = imnRdate_Year(1).Text & "/" & Format$(imnRdate_Month(1).Text, "00") & "/" & Format$(imnRdate_Day(1).Text, "00")
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF180"
    g_clsAdoAccess.Connection.Execute strSQL

    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF180"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '��t�f�[�^�I�[�v��
    strSQL = "SELECT vw_DT011.*,vw_DT010.Sname,vw_DT010.Addres,vw_DT010.Tel FROM vw_DT011"
    strSQL = strSQL & " LEFT OUTER JOIN vw_DT010 ON vw_DT010.Odate = vw_DT011.Odate AND vw_DT010.Pnum = vw_DT011.Pnum"
    strSQL = strSQL & " WHERE vw_DT011.Odate BETWEEN '" & strFromDate & "' AND '" & strToDate & "'"
    If Trim(txtIname.Text) <> "" Then
        strSQL = strSQL & " AND vw_DT011.Iname LIKE '%" & txtIname.Text & "%'"
    End If
    If Trim(txtSname.Text) <> "" Then
        strSQL = strSQL & " AND vw_DT010.Sname LIKE '%" & txtSname.Text & "%'"
    End If
    strSQL = strSQL & " ORDER BY vw_DT011.Odate,vw_DT011.Pnum"
    adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT011.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = adoDT011.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    
    Do While Not adoDT011.EOF
        wkRecordset.AddNew
        wkRecordset.Fields("Odate") = adoDT011.Fields("Odate")
        wkRecordset.Fields("Pnum") = adoDT011.Fields("Pnum")
        wkRecordset.Fields("Sname") = adoDT011.Fields("Sname")
        wkRecordset.Fields("Addres") = adoDT011.Fields("Addres")
        wkRecordset.Fields("Tel") = adoDT011.Fields("Tel")
        wkRecordset.Fields("Line") = adoDT011.Fields("Line")
        wkRecordset.Fields("Icode") = adoDT011.Fields("Icode")
        wkRecordset.Fields("Iname") = adoDT011.Fields("Iname")
        wkRecordset.Fields("Qty") = adoDT011.Fields("Qty")
            
        '�����f�[�^�I�[�v��
        strSQL = "SELECT * FROM vw_DT021" & _
                 " WHERE LEFT(Ocode,8) = '" & Format$(adoDT011.Fields("Odate"), "yyyymmdd") & "'" & _
                 " AND Pnum = " & adoDT011.Fields("Pnum") & _
                 " AND PnumLine = " & adoDT011.Fields("Line")
        adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoDT021.EOF = False Then
            wkRecordset.Fields("Price") = adoDT021.Fields("Price")
            wkRecordset.Fields("Bcode") = adoDT021.Fields("Bcode")
        End If
        adoDT021.Close
             
        wkRecordset.Update
             
        adoDT011.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    
    adoDT011.Close
    
    '�����f�[�^�I�[�v��
    strSQL = "SELECT vw_DT031.*,vw_DT030.Sname,vw_DT030.Addres,vw_DT030.Tel FROM vw_DT031"
    strSQL = strSQL & " LEFT OUTER JOIN vw_DT030 ON vw_DT030.Odate = vw_DT031.Odate AND vw_DT030.Onum = vw_DT031.Onum"
    strSQL = strSQL & " WHERE vw_DT031.Odate BETWEEN '" & strFromDate & "' AND '" & strToDate & "'"
    If Trim(txtIname.Text) <> "" Then
        strSQL = strSQL & " AND vw_DT031.Iname LIKE '%" & txtIname.Text & "%'"
    End If
    If Trim(txtSname.Text) <> "" Then
        strSQL = strSQL & " AND vw_DT030.Sname LIKE '%" & txtSname.Text & "%'"
    End If
    strSQL = strSQL & " ORDER BY vw_DT031.Odate,vw_DT031.Onum"
    adoDT031.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT031.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = adoDT031.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    
    Do While Not adoDT031.EOF
        wkRecordset.AddNew
        wkRecordset.Fields("Odate") = adoDT031.Fields("Odate")
        wkRecordset.Fields("Pnum") = adoDT031.Fields("Onum")
        wkRecordset.Fields("Sname") = adoDT031.Fields("Sname")
        wkRecordset.Fields("Addres") = adoDT031.Fields("Addres")
        wkRecordset.Fields("Tel") = adoDT031.Fields("Tel")
        wkRecordset.Fields("Line") = adoDT031.Fields("Line")
        wkRecordset.Fields("Icode") = adoDT031.Fields("Icode")
        wkRecordset.Fields("Iname") = adoDT031.Fields("Iname")
        wkRecordset.Fields("Qty") = adoDT031.Fields("Qty")
        wkRecordset.Fields("Price") = adoDT031.Fields("Price")
        wkRecordset.Fields("Bikou") = "*"
        wkRecordset.Update
             
        adoDT031.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    
    adoDT031.Close
    
    wkRecordset.Requery
    wkRecordset.Close
    
    MakePrintWork = True
    
MakePrintWork_Exit:
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    
    Exit Function

MakePrintWork_Cancel:

    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    MakePrintWork = False
    Call MsgBox("������[�N�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function

Private Sub imnRdate_Year_GotFocus(Index As Integer)

    imnRdate_Year(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Year_LostFocus(Index As Integer)

    imnRdate_Year(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRdate_Month_GotFocus(Index As Integer)

    imnRdate_Month(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Month_LostFocus(Index As Integer)

    imnRdate_Month(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRdate_Day_GotFocus(Index As Integer)

    imnRdate_Day(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Day_LostFocus(Index As Integer)

    imnRdate_Day(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    imnRdate_Year(0).SetFocus

End Sub

'�ځ@�I�@�@�FActiveReport�̈��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F0:�v���r���[ 1:���
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�T�^�O�X�^�P�R
'�X�V�����@�F
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt1 As New rptYpmf180_1
    Dim objRpt2 As New rptYpmf180_2
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    If Option1(0).Value = True Then
        With objArPrint
            .Name = "�o�i�ҐA�؈ꗗ�\"
            .objReport = objRpt1
            .Connection = g_clsAdoAccess.Connection
            .Caption = "�o�i�ҐA�؈ꗗ�\"
            If .PrintActiveReport(intFlg) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End With
    ElseIf Option1(1).Value = True Then
        With objArPrint
            .Name = "�A�؈ꗗ�\"
            .objReport = objRpt2
            .Connection = g_clsAdoAccess.Connection
            .Caption = "�A�؈ꗗ�\"
            If .PrintActiveReport(intFlg) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End With
    End If

    Set objRpt1 = Nothing
    Set objRpt2 = Nothing
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

Private Sub txtIname_GotFocus()

    txtIname.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtIname_LostFocus()

    txtIname.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtSname_GotFocus()

    txtSname.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtSname_LostFocus()

    txtSname.BackColor = FOCUS_NO_COLOR
    
End Sub

