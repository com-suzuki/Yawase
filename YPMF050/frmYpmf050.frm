VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmYpmf050 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   4320
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
   Icon            =   "frmYpmf050.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   12150
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdLogin 
         Caption         =   "�J�ÔN�����ƒS���҂̕ύX"
         Height          =   375
         Left            =   6960
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9900
         TabIndex        =   18
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
         Contents        =   "frmYpmf050.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   15
      Top             =   3540
      Width           =   12015
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   9
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
         Picture         =   "frmYpmf050.frx":0D13
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10200
         TabIndex        =   11
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
         Picture         =   "frmYpmf050.frx":0D2F
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   8460
         TabIndex        =   10
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
         Picture         =   "frmYpmf050.frx":0E89
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   2895
      Left            =   60
      TabIndex        =   14
      Top             =   660
      Width           =   12015
      Begin VB.CheckBox chkMishu 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8100
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2160
         Width           =   555
      End
      Begin VB.CheckBox chkIji 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame fraReprint 
         BorderStyle     =   0  '�Ȃ�
         Height          =   675
         Left            =   2340
         TabIndex        =   30
         Top             =   2100
         Width           =   3435
         Begin imNumber6Ctl.imNumber imnRePrintNum 
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   180
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   767
            Calculator      =   "frmYpmf050.frx":0F9B
            Caption         =   "frmYpmf050.frx":0FBB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf050.frx":1029
            Keys            =   "frmYpmf050.frx":1047
            Spin            =   "frmYpmf050.frx":1091
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
            ValueVT         =   2011365381
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRePrintNum 
            Height          =   435
            Index           =   1
            Left            =   1980
            TabIndex        =   7
            Top             =   180
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   767
            Calculator      =   "frmYpmf050.frx":10B9
            Caption         =   "frmYpmf050.frx":10D9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf050.frx":1147
            Keys            =   "frmYpmf050.frx":1165
            Spin            =   "frmYpmf050.frx":11AF
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
         Begin VB.Label lblRePrintNum 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "���"
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
            Index           =   1
            Left            =   2640
            TabIndex        =   32
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblRePrintNum 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "��� �`"
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
            Index           =   0
            Left            =   780
            TabIndex        =   31
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  '�Ȃ�
         Height          =   615
         Left            =   1620
         TabIndex        =   29
         Top             =   1500
         Width           =   3015
         Begin VB.OptionButton optFdiv 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   0
            Style           =   1  '���̨���
            TabIndex        =   3
            Top             =   120
            Width           =   1395
         End
         Begin VB.OptionButton optFdiv 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   1440
            Style           =   1  '���̨���
            TabIndex        =   4
            Top             =   120
            Width           =   1395
         End
      End
      Begin VB.CheckBox chkRePrint 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2160
         Width           =   555
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1620
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
      Begin CSComboLib.CSComboBox cboBcode 
         Height          =   405
         Index           =   0
         Left            =   1620
         TabIndex        =   1
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
         Contents        =   "frmYpmf050.frx":11D7
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   2220
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�Ĕ��s"
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
      Begin CSComboLib.CSComboBox cboBcode 
         Height          =   405
         Index           =   1
         Left            =   1620
         TabIndex        =   2
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
         Contents        =   "frmYpmf050.frx":11F0
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   6000
         TabIndex        =   34
         Top             =   1620
         Visible         =   0   'False
         Width           =   4110
         _Version        =   262145
         _ExtentX        =   7250
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�Q��ڈȍ~�͈ێ��Ǘ���𒥎����Ȃ�"
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
         LabelWidth      =   241
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
         Index           =   3
         Left            =   6000
         TabIndex        =   35
         Top             =   2220
         Width           =   2010
         _Version        =   262145
         _ExtentX        =   3545
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "����ү���ޕ\��"
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
         LabelWidth      =   107
         LabelHeight     =   25
         LabelLeft       =   13
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
         TabIndex        =   28
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblScode_Name 
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
         TabIndex        =   27
         Top             =   1020
         Width           =   9195
      End
      Begin VB.Label lblScode_Name 
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
         TabIndex        =   26
         Top             =   180
         Width           =   9195
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
      Caption         =   "frmYpmf050.frx":1209
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf050.frx":1277
      Key             =   "frmYpmf050.frx":1295
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
      TabIndex        =   12
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf050.frx":12D9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf050.frx":1347
      Key             =   "frmYpmf050.frx":1365
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
      TabIndex        =   13
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf050.frx":13A9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf050.frx":1417
      Key             =   "frmYpmf050.frx":1435
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
Attribute VB_Name = "frmYpmf050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CHUMON_ONLY_FLG = "1"          '1:�����݂̂̔���͎萔�����Ƃ�Ȃ� 0:�萔���Ƃ�

Private curBrate2 As Integer        '201107 �萔���Q

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
    
    If Trim(cboBcode(Index).Text) = "" Then
        lblScode_Name(Index).Caption = ""
        Exit Sub
    End If
    If IsNumeric(cboBcode(Index).Text) = False Then
        cboBcode(Index).Text = ""
        lblScode_Name(Index).Caption = ""
        Exit Sub
    End If
    If cboBcode(Index).Tag = cboBcode(Index).Text Then Exit Sub
    
    lblScode_Name(Index).Caption = ""
    
    With adoRecordset1
        '���Ӑ�}�X�^
        strSQL = "{call sp_MT070;2(" & Trim(cboBcode(Index).Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Fdiv")) Then
                If .Fields("Fdiv") = BUSINESS_DIV_BUYER Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    lblScode_Name(Index).Caption = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If Index = 0 Then
        cboBcode(1).Text = cboBcode(0).Text
        lblScode_Name(1).Caption = lblScode_Name(0).Caption
    End If
    
    Exit Sub

cboBcode_Validate_Err:

    Call MsgBox("�t�H�[�J�X�ړ��O�G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBcode_Validate_Err")

End Sub

Private Sub chkRePrint_Click()
    
    On Error Resume Next
    
    If chkRePrint.Value = 1 Then
        fraReprint.Visible = True
        imnRePrintNum(0).SetFocus
    Else
        fraReprint.Visible = False
    End If
    
End Sub

'�ځ@�I�@�@�F
'���@���@�@�F��ʃN���A�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
'�X�V�����@�F
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    cboBcode(0).SetFocus

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
'�X�V�����@�F
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If MsgBox("���s���܂����H", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

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
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
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
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
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
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

'    Me.Caption = SYSTEM_NAME & "-" & "���吸�Z"
    Me.Caption = "���吸�Z"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    Call FieldsClear(0)
    
    If Command() <> "" Then
        '���̓`�F�b�N
        If DoValidationChecks() = False Then End
        '����p���[�N�쐬
        If MakePrintWork() = False Then End
        '���
        If ActiveReportPrint(0) = False Then End
        End
    End If
    
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
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
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
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
'�X�V�����@�F
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        '2005/09/01 �C��
        If Trim$(g_strBcode) = "" Then
            cboBcode(0).Text = ""
            cboBcode(0).Tag = ""
            cboBcode(1).Text = ""
            cboBcode(1).Tag = ""
        Else
            cboBcode(0).Text = g_strBcode
            cboBcode(0).Tag = g_strBcode
            cboBcode(1).Text = g_strBcode
            cboBcode(1).Tag = g_strBcode
        End If
        
        lblScode_Name(0).Caption = ""
        lblScode_Name(1).Caption = ""
        optFdiv(0).Value = True
        
        If Trim$(g_strRePrintNum) = "0" Then
            chkRePrint.Value = 0
            fraReprint.Visible = False
            imnRePrintNum(0).Value = 0
            imnRePrintNum(1).Value = 99
        Else
            chkRePrint.Value = 1
            fraReprint.Visible = True
            imnRePrintNum(0).Value = g_strRePrintNum
            imnRePrintNum(1).Value = g_strRePrintNum
        End If
        
        chkIji.Value = 1
        chkMishu.Value = 1
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
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
'�X�V�����@�F
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "�J�ÔN��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
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
'    If chkRePrint.Value = 1 Then
'        If imnRePrintNum(0).Value = 0 Or imnRePrintNum(1).Value = 0 Then
'            imnRePrintNum(0).SetFocus
'            strErrMsg = "�񐔂���͂��Ă��������B"
'            GoTo ErrorTrap:
'        End If
'    End If
    
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
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
'�X�V�����@�F�Q�O�O�S�^�O�P�^�R�P�@�����݂̂̏ꍇ�͎萔�����Ƃ�Ȃ�
'�@�@�@�@�@�@�Q�O�O�T�^�O�W�^�P�Q�@�������̏ڍׂ��󎚂���
'�@�@�@�@�@�@�Q�O�O�U�^�O�U�^�O�W�@����o�^����ǉ�
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim adoDT010 As New ADODB.Recordset
    Dim adoDT021 As New ADODB.Recordset
    Dim adoDT021M As New ADODB.Recordset
    Dim adoDT030 As New ADODB.Recordset
    Dim adoDT031M As New ADODB.Recordset
    Dim adoDT041 As New ADODB.Recordset
    Dim adoDT041TEMP As New ADODB.Recordset
    Dim adoDT060 As New ADODB.Recordset
    Dim adoDT060TEMP As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim wkRecordsetTemp As New ADODB.Recordset
    Dim strBuff1 As String
    
    Dim intLine As Integer                  '�s�ԍ�
    Dim lngCount As Long                    '���R�[�h����
    Dim curBkeep As Currency                '����ێ��Ǘ���(�W��)
    Dim curBkeepCurrent As Currency         '����ێ��Ǘ���(����)
    Dim intBfraction As Integer             '����[������
    Dim intNum As Integer                   '��
    Dim curTaxRate As Currency              '����ŗ�
    Dim intR As Integer                     '�������
    Dim curBRounding As Currency            '����ۂߒP��
    Dim strMemo As String                   '�`�[�̃���
    Dim strOdateNum As String               '�J�Ó�(YYYYMMDD�`��)

    Dim curBrate2Current As Currency        '201107 �萔���Q(����)
    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
'********** �������� **********
    
    strOdateNum = Global_Get_NumericDay(lblOdate.Caption)
    
    '������
    lngCount = 0        '���R�[�h����
    curBkeep = 0        '����ێ��Ǘ���
    intBfraction = 0    '����[������
    curTaxRate = 0      '����ŗ�
    intR = 0            '�������
    curBRounding = 0    '����ۂߒP��
    strMemo = ""
    
    curBrate2 = 0       '201107 �����萔��
    
    If optFdiv(0).Value = True Then
        intR = PAYMENT_DIV_CASH
    ElseIf optFdiv(1).Value = True Then
        intR = PAYMENT_DIV_TRANSFER
    End If
    
    '�ݒ�}�X�^�I�[�v��
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        If Not IsNull(adoMT010.Fields("Bkeep")) Then curBkeep = adoMT010.Fields("Bkeep")
        If Not IsNull(adoMT010.Fields("Bfraction")) Then intBfraction = adoMT010.Fields("Bfraction")
        If Not IsNull(adoMT010.Fields("BRounding")) Then curBRounding = adoMT010.Fields("BRounding")
        If Not IsNull(adoMT010.Fields("Memo")) Then strMemo = adoMT010.Fields("Memo")
        If Not IsNull(adoMT010.Fields("Brate2")) Then curBrate2 = adoMT010.Fields("Brate2")
    End If
    adoMT010.Close
    
    '����ŗ��擾
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, Trim(lblOdate.Caption))
    
'********** ���[�N **********
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF050"
    g_clsAdoAccess.Connection.Execute strSQL

    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF050"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
'********** �����f�[�^ **********
    
    '�����f�[�^�I�[�v��
    If chkRePrint.Value = 0 Then
        strSQL = "{call sp_YPMF0501;1('" & strOdateNum & "'," & _
                  cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    Else
        strSQL = "{call sp_YPMF0501;2('" & strOdateNum & "'," & _
                  cboBcode(0).Text & "," & cboBcode(1).Text & "," & _
                  imnRePrintNum(0).Text & "," & imnRePrintNum(1).Text & ")}"
    End If
    adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT021.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = adoDT021.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    
    Do While Not adoDT021.EOF
        intLine = 1     '�s�ԍ�
        
        '�������׃f�[�^�I�[�v��
        If chkRePrint.Value = 0 Then
            strSQL = "{call sp_YPMF0502;1('" & strOdateNum & "'," & _
                      adoDT021.Fields("Bcode") & "," & adoDT021.Fields("Bcode") & ")}"
        Else
            strSQL = "{call sp_YPMF0502;2('" & strOdateNum & "'," & _
                      adoDT021.Fields("Bcode") & "," & adoDT021.Fields("Bcode") & "," & _
                      adoDT021.Fields("Bnum") & "," & adoDT021.Fields("Bnum") & ")}"
        End If
        adoDT021M.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        Do While Not adoDT021M.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Key1") = Format(adoDT021M.Fields("Bcode"), "0000") & Format(adoDT021M.Fields("Bnum"), "00")
            wkRecordset.Fields("Key2") = Format(adoDT021M.Fields("Bcode"), "0000") & Format(adoDT021M.Fields("Bnum"), "00")
            wkRecordset.Fields("Num") = adoDT021M.Fields("Bnum")
            wkRecordset.Fields("Div") = "A"
            wkRecordset.Fields("Odate") = Trim(lblOdate.Caption)
            wkRecordset.Fields("Bcode") = adoDT021M.Fields("Bcode")
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, adoDT021M.Fields("Bcode"), Trim(lblOdate.Caption), strBuff1)
            wkRecordset.Fields("Icode") = adoDT021M.Fields("Icode")
            wkRecordset.Fields("Iname") = adoDT021M.Fields("Iname")
            wkRecordset.Fields("Qty") = adoDT021M.Fields("Qty")
            wkRecordset.Fields("Price1") = adoDT021M.Fields("Price")
            
            '��t�f�[�^
            strSQL = "{call sp_YPMF0503('" & Trim(lblOdate.Caption) & "'," & adoDT021M.Fields("Pnum") & "," & adoDT021M.Fields("PnumLine") & ")}"
            adoDT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT010.EOF = False Then
                wkRecordset.Fields("Pnum") = adoDT010.Fields("Pnum")
                wkRecordset.Fields("Sname") = adoDT010.Fields("Sname")
            End If
            adoDT010.Close
'            wkRecordset.Fields("Total") = 0
'            wkRecordset.Fields("Tax") = 0
            wkRecordset.Fields("Keep") = curBkeep
            wkRecordset.Fields("Brate2") = curBrate2
'            wkRecordset.Fields("GTotal") = 0
            wkRecordset.Fields("Itime") = Format(Now(), "hh��mm��")
            wkRecordset.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
            wkRecordset.Fields("Pname") = Trim(lblPname.Caption)
            If Not IsNull(adoDT021M.Fields("Sline")) Then
                wkRecordset.Fields("Idiv") = adoDT021M.Fields("Sline")
            Else
                wkRecordset.Fields("Idiv") = 0
            End If
            wkRecordset.Fields("Ocode") = right$(adoDT021M.Fields("Ocode"), 4)
            wkRecordset.Fields("Memo") = strMemo
            wkRecordset.Update
        
            If adoDT021M.Fields("Bnum") = 0 Then
                '���吸�Z�f�[�^���琸�Z�񐔂��擾
                strSQL = "SELECT * FROM DT041" & _
                         " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                         " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                         " ORDER BY Odate,Bcode,Num DESC"
                adoDT041.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                intNum = 1
                If adoDT041.EOF = False Then
                    intNum = CInt(adoDT041.Fields("Num")) + 1
                End If
                adoDT041.Close
                
                '�������׃f�[�^�X�V
                adoDT021M.Fields("Bdiv") = BUYER_REPORT_ON
                adoDT021M.Fields("Bnum") = intNum
                adoDT021M.Update
            End If
    
            adoDT021M.MoveNext
            intLine = intLine + 1
        Loop
        adoDT021M.Close
        
        adoDT021.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    adoDT021.Close
    
'********** �����f�[�^�擾 **********
    
    '�����f�[�^�I�[�v��
    If chkRePrint.Value = 0 Then
        strSQL = "{call sp_YPMF0504;1('" & lblOdate.Caption & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    Else
        strSQL = "{call sp_YPMF0504;2('" & lblOdate.Caption & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & "," & _
                  imnRePrintNum(0).Text & "," & imnRePrintNum(1).Text & ")}"
    End If
    adoDT030.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoDT030.EOF
        '�������׃f�[�^�I�[�v��
        If chkRePrint.Value = 0 Then
            strSQL = "{call sp_YPMF0505;1('" & lblOdate.Caption & "'," & adoDT030.Fields("Bcode") & ")}"
        Else
            strSQL = "{call sp_YPMF0505;2('" & lblOdate.Caption & "'," & adoDT030.Fields("Bcode") & "," & adoDT030.Fields("Bnum") & ")}"
        End If
        adoDT031M.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If adoDT031M.EOF = False Then
            '�J�n�s�ԍ��擾
            intLine = 1

            '���[�N�I�[�v��
            strSQL = "SELECT WK_YPMF050.Odate, WK_YPMF050.Bcode, WK_YPMF050.Line" & _
                     " FROM WK_YPMF050 " & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Bcode = " & adoDT031M.Fields("Bcode") & _
                     " AND Num = " & adoDT031M.Fields("Bnum") & _
                     " ORDER BY WK_YPMF050.Odate, WK_YPMF050.Bcode, WK_YPMF050.Line DESC"
            wkRecordsetTemp.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
            If wkRecordsetTemp.EOF = False Then
                If Not IsNull(wkRecordsetTemp.Fields("Line")) Then
                    intLine = CInt(wkRecordsetTemp.Fields("Line")) + 1
                End If
            End If
            wkRecordsetTemp.Close
        End If
        Do While Not adoDT031M.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Key1") = Format(adoDT031M.Fields("Bcode"), "0000") & Format(adoDT031M.Fields("Bnum"), "00")
            wkRecordset.Fields("Key2") = Format(adoDT031M.Fields("Bcode"), "0000") & Format(adoDT031M.Fields("Bnum"), "00")
            wkRecordset.Fields("Num") = adoDT031M.Fields("Bnum")
            wkRecordset.Fields("Div") = "B"
            wkRecordset.Fields("Odate") = Trim(lblOdate.Caption)
            wkRecordset.Fields("Bcode") = adoDT031M.Fields("Bcode")
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, adoDT031M.Fields("Bcode"), Trim(lblOdate.Caption), strBuff1)
            wkRecordset.Fields("Icode") = adoDT031M.Fields("Icode")
            wkRecordset.Fields("Iname") = adoDT031M.Fields("Iname")
            wkRecordset.Fields("Qty") = adoDT031M.Fields("Qty")
            wkRecordset.Fields("Price1") = adoDT031M.Fields("Price")
            wkRecordset.Fields("Pnum") = adoDT031M.Fields("Onum")
            wkRecordset.Fields("Sname") = adoDT031M.Fields("Sname")
'            wkRecordset.Fields("Total") = 0
'            wkRecordset.Fields("Tax") = 0
            wkRecordset.Fields("Keep") = curBkeep
'            wkRecordset.Fields("GTotal") = 0
            wkRecordset.Fields("Itime") = Format(Now(), "hh��mm��")
            wkRecordset.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
            wkRecordset.Fields("Pname") = Trim(lblPname.Caption)
            wkRecordset.Fields("Idiv") = 0
            wkRecordset.Fields("Ocode") = Format$(adoDT031M.Fields("Onum"), "0000") & "*"
            wkRecordset.Fields("Memo") = strMemo
            wkRecordset.Update

            If adoDT031M.Fields("Bnum") = 0 Then
                '���吸�Z�f�[�^���琸�Z�񐔂��擾
                strSQL = "SELECT * FROM DT041" & _
                         " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                         " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                         " ORDER BY Odate,Bcode,Num DESC"
                adoDT041.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                intNum = 1
                If adoDT041.EOF = False Then
                    intNum = CInt(adoDT041.Fields("Num")) + 1
                End If
                adoDT041.Close

                '�������׃f�[�^�X�V
                adoDT031M.Fields("Bdiv") = BUYER_REPORT_ON
                adoDT031M.Fields("Bnum") = intNum
                adoDT031M.Update
            End If

            adoDT031M.MoveNext
            intLine = intLine + 1
        Loop
        adoDT031M.Close

        adoDT030.MoveNext
    Loop
    adoDT030.Close

    wkRecordset.Close
    
'********** ���吸�Z�f�[�^�쐬 **********

    '���[�N�I�[�v��
    'strSQL = "SELECT WK_YPMF050.Odate, WK_YPMF050.Bcode, WK_YPMF050.Key1, WK_YPMF050.Num, Sum(WK_YPMF050.Price1) AS Price1_Total" & _
    '         " FROM WK_YPMF050 " & _
    '         " GROUP BY WK_YPMF050.Odate, WK_YPMF050.Bcode, WK_YPMF050.Key1, WK_YPMF050.Num" & _
    '         " ORDER BY WK_YPMF050.Odate, WK_YPMF050.Bcode, WK_YPMF050.Key1, WK_YPMF050.Num"
    
    '201107 Price2_Total�Ƃ��ċ������̍��v�z���W�v...�������͏���
    strSQL = "SELECT M.Odate,M.Bcode,M.Key1,M.Num,Sum(M.Price1) AS Price1_Total," & _
            "(SELECT Sum(S.Price1) FROM  WK_YPMF050 S " & _
            "WHERE Div='A' AND M.Odate=S.Odate AND M.Bcode=S.Bcode AND M.Key1=S.Key1 AND M.Num=S.Num " & _
            "GROUP BY S.Odate, S.Bcode, S.Key1, S.Num ) AS Price2_Total " & _
            "FROM WK_YPMF050 M " & _
            "GROUP BY M.Odate, M.Bcode, M.Key1, M.Num " & _
            "ORDER BY M.Odate, M.Bcode, M.Key1, M.Num"
             
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    If wkRecordset.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = wkRecordset.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    
    lngCount = 0
    Do While Not wkRecordset.EOF
    
        If wkRecordset.Fields("Num") = 0 Then
            '���吸�Z�f�[�^�I�[�v��
            strSQL = "SELECT * FROM DT041" & _
                     " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                     " ORDER BY Odate,Bcode,Num DESC"
            adoDT041.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If adoDT041.EOF = False Then
                intNum = CInt(adoDT041.Fields("Num")) + 1
            Else
                intNum = 1
            End If
            adoDT041.AddNew
        Else
            intNum = wkRecordset.Fields("Num")
            
            '���吸�Z�f�[�^�I�[�v��
            strSQL = "SELECT * FROM DT041" & _
                     " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                     " AND Num = " & intNum & _
                     " ORDER BY Odate,Bcode,Num DESC"
            adoDT041.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If adoDT041.EOF = True Then
                adoDT041.AddNew
            End If
        End If
        
'        '�Q��ڈȍ~�̐��Z�̏ꍇ�͎萔�������Ȃ�
'        curBkeepCurrent = curBkeep
'        If intNum >= 2 And chkIji.Value = 1 Then
'            curBkeepCurrent = 0
'        End If
        
        '�����݂̂̏ꍇ�́A�萔�������Ȃ�(2004/01/31�ǉ�)
        If CHUMON_ONLY_FLG = "1" Then
            curBkeepCurrent = curBkeep
            curBrate2Current = curBrate2 '201107

            If chkIji.Value = 1 Then
                
                '����
                If intNum = 1 Then
                    '�����f�[�^���Ȃ��ꍇ�͎萔�������Ȃ�
                    strSQL = "SELECT * FROM DT021" & _
                             " WHERE LEFT(Ocode, 8) = '" & strOdateNum & "'" & _
                             " AND Bcode = " & wkRecordset.Fields("Bcode")
                    adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                    If adoDT021.EOF = True Then
                        curBkeepCurrent = 0
                        curBrate2Current = 0   '2011/07
                        adoDT021.Close
                    Else
                        adoDT021.Close
                        
                        '�����f�[�^�����邪�A���Z�񐔂��Q��ڈȍ~�̏ꍇ�͎��Ȃ�
                        strSQL = "SELECT * FROM DT021" & _
                                 " WHERE LEFT(Ocode, 8) = '" & strOdateNum & "'" & _
                                 " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                                 " AND Bnum = 1 "
                        adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                        If adoDT021.EOF = False Then
                            adoDT021.Close
                        Else
                            adoDT021.Close

                            strSQL = "SELECT * FROM DT021" & _
                                     " WHERE LEFT(Ocode, 8) = '" & strOdateNum & "'" & _
                                     " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                                     " AND Bnum >= 2 "
                            adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                            If adoDT021.EOF = False Then
                                curBkeepCurrent = 0
                                'curBrate2Current = 0   2011/11/07 �����萔���͑ΏۊO
                            End If
                            adoDT021.Close
                        End If
                    End If
                    
                End If
            
                '�Q��ڈȍ~
                If intNum >= 2 Then
'2004/10/08 �Q��ڈȍ~�́A�����f�[�^�̗L���Ɋ֌W�Ȃ��ێ��Ǘ��萔���̓[���ɂ���

'                    '�����f�[�^��T��
'                    strSQL = "SELECT * FROM DT021" & _
'                             " WHERE LEFT(Ocode, 8) = '" & strOdateNum & "'" & _
'                             " AND Bcode = " & wkRecordset.Fields("Bcode") & _
'                             " AND Bnum < " & intNum
'                    adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
'                    If adoDT021.EOF = False Then
'                        curBkeepCurrent = 0
'                    End If
'                    adoDT021.Close
                    
'2005/09/15 �Q��ڈȍ~�́A�P��ڂɋ����f�[�^���Ȃ��ꍇ�͈ێ��Ǘ��萔�����Ƃ�
                    
                    '�����f�[�^��T��
                    strSQL = "SELECT * FROM DT021" & _
                             " WHERE LEFT(Ocode, 8) = '" & strOdateNum & "'" & _
                             " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                             " AND Bnum < " & intNum
                    adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                    If adoDT021.EOF = False Then
                        curBkeepCurrent = 0
                        'curBrate2Current = 0   2011/11/07 �����萔���͑ΏۊO
                    End If
                    adoDT021.Close
                End If
            
            End If
            
        Else
            '�Q��ڈȍ~�̐��Z�̏ꍇ�͎萔�������Ȃ�
            curBkeepCurrent = curBkeep
            curBrate2Current = curBrate2
            
            If intNum >= 2 And chkIji.Value = 1 Then
                curBkeepCurrent = 0
                'curBrate2Current = 0   2011/11/07 �����萔���͑ΏۊO
            End If
        End If
        
        
        adoDT041.Fields("Odate") = wkRecordset.Fields("Odate")
        adoDT041.Fields("Bcode") = wkRecordset.Fields("Bcode")
        adoDT041.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, wkRecordset.Fields("Bcode"), Trim(lblOdate.Caption), strBuff1)
        adoDT041.Fields("Num") = intNum
        adoDT041.Fields("Total") = wkRecordset.Fields("Price1_Total")
        
        'adoDT041.Fields("Tax") = Global_Get_Tax(adoDT041.Fields("Total"), curTaxRate, intBfraction, curBRounding)
        
        '201107
        'adoDT041.Fields("GTotal") = CCur(adoDT041.Fields("Total")) + CCur(adoDT041.Fields("Tax")) + CCur(adoDT041.Fields("Keep"))
        '201107 �������̍��v�z����萔�����v�Z
        If IsNull(wkRecordset.Fields("Price2_Total")) Then
        adoDT041.Fields("Brate2") = 0
        Else
        adoDT041.Fields("Brate2") = Global_Get_Brate(Global_Rounding(wkRecordset.Fields("Price2_Total") / (1 + (curTaxRate / 100)), intBfraction, 1), curBrate2Current, intBfraction, 1)
        End If
        'adoDT041.Fields("GTotal") = CCur(adoDT041.Fields("Total")) + CCur(adoDT041.Fields("Tax")) + CCur(adoDT041.Fields("Keep")) + CCur(adoDT041.Fields("Brate2"))
        '201107
        
        '202308 �ێ��Ǘ���
        adoDT041.Fields("Keep") = Global_Rounding(curBkeepCurrent / (1 + (curTaxRate / 100)), intBfraction, 1)
        '202308 ����Ōv�Z�@�������z�{�������z-�萔��-�Ŕ����ێ��Ǘ���@����v�Z
        adoDT041.Fields("Tax") = Global_Get_Tax((adoDT041.Fields("Total") + CCur(adoDT041.Fields("Keep")) + adoDT041.Fields("Brate2")), curTaxRate, intBfraction, 1)
        adoDT041.Fields("GTotal") = CCur(adoDT041.Fields("Total")) + CCur(adoDT041.Fields("Keep")) + CCur(adoDT041.Fields("Brate2") + CCur(adoDT041.Fields("Tax")))
        
        
        If optFdiv(0).Value = True Then
            adoDT041.Fields("Rdiv") = PAYMENT_ON
            adoDT041.Fields("Rdate") = Format(Now(), "yyyy/mm/dd")
            adoDT041.Fields("R") = intR
        Else
            adoDT041.Fields("Rdiv") = PAYMENT_OFF
            adoDT041.Fields("Rdate") = Null
            adoDT041.Fields("R") = Null
        End If
        adoDT041.Fields("Itime") = Format(Now(), "hh:mm:ss")
        adoDT041.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
        
        
        adoDT041.Update
    
        '���[�N�̍��v�z�Ȃǂ��X�V
        strSQL = "UPDATE WK_YPMF050" & _
                 " SET WK_YPMF050.Total = " & adoDT041.Fields("Total") & "," & _
                 " WK_YPMF050.Tax = " & adoDT041.Fields("Tax") & "," & _
                 " WK_YPMF050.Keep = " & adoDT041.Fields("Keep") & "," & _
                 " WK_YPMF050.Brate2 = " & adoDT041.Fields("Brate2") & "," & _
                 " WK_YPMF050.GTotal = " & adoDT041.Fields("GTotal") & _
                 " WHERE WK_YPMF050.Odate = '" & adoDT041.Fields("Odate") & "'" & _
                 " AND WK_YPMF050.Bcode = " & adoDT041.Fields("Bcode") & _
                 " AND WK_YPMF050.Key1 = '" & wkRecordset.Fields("Key1") & "'" & _
                 " AND WK_YPMF050.Num = " & wkRecordset.Fields("Num")
        g_clsAdoAccess.Connection.Execute strSQL
        
        '********** ������ **********
        If chkMishu.Value = 1 Then
            
            '���吸�Z�f�[�^�I�[�v��
'            strSQL = "SELECT * FROM DT041" & _
                     " WHERE Odate < '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                     " AND Rdiv = " & PAYMENT_OFF & _
                     " ORDER BY Odate,Bcode,Num"
            strSQL = "SELECT Odate,Bcode,SUM(DT041.GTotal) AS GTotal FROM DT041" & _
                     " WHERE Odate < '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND Bcode = " & wkRecordset.Fields("Bcode") & _
                     " AND Rdiv = " & PAYMENT_OFF & _
                     " GROUP BY Odate,Bcode" & _
                     " ORDER BY Odate"
            adoDT041TEMP.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT041TEMP.EOF = False Then
                
                '2005/08/12 �����̖��ו\���ǉ�
                Dim curZandaka As Currency
                Do While Not adoDT041TEMP.EOF
                    If IsNull(adoDT041TEMP.Fields("GTotal")) = False Then
                        curZandaka = adoDT041TEMP.Fields("GTotal")
                    Else
                        curZandaka = 0
                    End If
                
                    '�����f�[�^��������z�������Ďc���v�Z
                    strSQL = "SELECT Ptotal FROM DT060" & _
                             " WHERE Odate = '" & adoDT041TEMP("Odate") & "'" & _
                             " AND Bcode = " & adoDT041TEMP("Bcode")
                    adoDT060TEMP.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                    Do While Not adoDT060TEMP.EOF
                        curZandaka = curZandaka - CCur(adoDT060TEMP.Fields("Ptotal"))
                        
                        adoDT060TEMP.MoveNext
                    Loop
                    adoDT060TEMP.Close
                    
                    If curZandaka > 0 Then
                        strSQL = "INSERT INTO WK_YPMF050(Key1,Key2,Num,Odate,Bcode,Line,Iname,Price1,Mishu)"
                        strSQL = strSQL & "VALUES('" & wkRecordset.Fields("Key1") & "',"
                        strSQL = strSQL & "'" & wkRecordset.Fields("Key1") & "',"
                        strSQL = strSQL & wkRecordset.Fields("Num") & ","
                        strSQL = strSQL & "'" & wkRecordset.Fields("Odate") & "',"
                        strSQL = strSQL & wkRecordset.Fields("Bcode") & ","
                        strSQL = strSQL & "0" & ","
                        strSQL = strSQL & "'�������@�@�J�Ó�:" & adoDT041TEMP("Odate") & "',"
                        strSQL = strSQL & curZandaka & ","
                        strSQL = strSQL & "1" & ")"
                        g_clsAdoAccess.Connection.Execute strSQL
                    End If
                    
                    adoDT041TEMP.MoveNext
                Loop
                
                adoDT041TEMP.MoveFirst
                
                '�����̃t���O���X�V
                strSQL = "UPDATE WK_YPMF050" & _
                         " SET WK_YPMF050.CFlg = 1" & _
                         " WHERE WK_YPMF050.Odate = '" & adoDT041.Fields("Odate") & "'" & _
                         " AND WK_YPMF050.Bcode = " & adoDT041.Fields("Bcode")
                g_clsAdoAccess.Connection.Execute strSQL
            
            
            End If
            adoDT041TEMP.Close
        End If
        
        adoDT041.Close
        wkRecordset.MoveNext
        lngCount = lngCount + 1
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    
    wkRecordset.Requery
    wkRecordset.Close
    
    
'********** 2006/06/08 ���Ӑ�}�X�^�̓o�^����ǉ� **********
    
    Dim adoMT070 As New ADODB.Recordset
    
    strSQL = "SELECT Bcode FROM WK_YPMF050 WHERE Bcode IS NOT NULL GROUP BY Bcode"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockReadOnly
    If wkRecordset.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = wkRecordset.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    Do While Not wkRecordset.EOF
        If IsNull(wkRecordset("Bcode")) = False Then
            strSQL = "SELECT Adddate FROM MT070 WHERE Bcode = " & wkRecordset("Bcode")
            adoMT070.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoMT070.EOF = False Then
                If IsNull(adoMT070("Adddate").Value) = False Then
                    strSQL = " UPDATE WK_YPMF050 SET Adddate = #" & Format$(adoMT070("Adddate").Value, "yyyy/mm/dd") & "#"
                    strSQL = strSQL & " WHERE Bcode = " & wkRecordset("Bcode")
                    g_clsAdoAccess.Connection.Execute strSQL
                End If
            End If
            adoMT070.Close
        End If
        wkRecordset.MoveNext
    
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    wkRecordset.Close
    
    '�f�[�^�������������f����Ȃ��\�������邽��(�o�O?)���N�G���[���Ă���
    strSQL = "SELECT * FROM WK_YPMF050"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockReadOnly
    wkRecordset.Requery
    wkRecordset.Close
    
'********************************************************
    
    g_clsAdoSQL.Connection.CommitTrans
    
    If lngCount = 0 Then
        Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If
    
    MakePrintWork = True
    
MakePrintWork_Exit:
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    
    Exit Function

MakePrintWork_Cancel:

    g_clsAdoSQL.Connection.RollbackTrans
    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    MakePrintWork = False
    Call MsgBox("������[�N�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function

'�ځ@�I�@�@�F�R���{�{�b�N�X�̍쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
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

Private Sub imnRePrintNum_GotFocus(Index As Integer)

    imnRePrintNum(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRePrintNum_LostFocus(Index As Integer)

    imnRePrintNum(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cboBcode(0).SetFocus

End Sub

Private Sub optFdiv_Click(Index As Integer)

    Dim intIndex1 As Integer

    On Error GoTo optFdiv_Click_Err

    '�w�i�F�̕ύX
    For intIndex1 = 0 To optFdiv.Count - 1
        If intIndex1 = Index Then
            optFdiv(intIndex1).BackColor = BUTTON_ON
        Else
            optFdiv(intIndex1).BackColor = BUTTON_OFF
        End If
    Next intIndex1

    Exit Sub

optFdiv_Click_Err:

    Call MsgBox("������ʃN���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "optFdiv_Click_Err")

End Sub

'�ځ@�I�@�@�FActiveReport�̈��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F0:�v���r���[ 1:���
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
'�X�V�����@�F
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf050
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
'    objRpt.lblRate.Caption = "(" & curBrate2 & "%)"
'
    With objArPrint
        .Name = "���吸�Z���ו[���̎���"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .Caption = "���吸�Z���ו[���̎���"
        
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
    Call MsgBox("���s�N���b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReportPrint_Err")
    
End Function

