VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmYpmf100 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   10395
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf100.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   15000
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   9840
      TabIndex        =   162
      Top             =   0
      Width           =   2895
      Begin VB.CheckBox chkFlg 
         Caption         =   "���������\��"
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
         Left            =   120
         TabIndex        =   163
         Top             =   120
         Width           =   2595
      End
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   60
      TabIndex        =   157
      Top             =   660
      Width           =   14835
      Begin VB.CommandButton cmdSearch 
         Caption         =   "�����J�n"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13080
         TabIndex        =   3
         Top             =   180
         Width           =   1635
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
         Contents        =   "frmYpmf100.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   158
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "���庰�ޒ��o"
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
         LabelWidth      =   87
         LabelHeight     =   25
         LabelLeft       =   5
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
         Left            =   7620
         TabIndex        =   2
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
         Contents        =   "frmYpmf100.frx":0D13
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
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
         TabIndex        =   161
         Top             =   180
         Width           =   4395
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
         Left            =   8700
         TabIndex        =   160
         Top             =   180
         Width           =   4275
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
         Left            =   7140
         TabIndex        =   159
         Top             =   180
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1980
      Top             =   9840
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   126
      Top             =   9000
      Width           =   14835
      Begin imNumber6Ctl.imNumber imnTotal_Total 
         Height          =   375
         Left            =   5340
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   180
         Width           =   1300
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   661
         Calculator      =   "frmYpmf100.frx":0D2C
         Caption         =   "frmYpmf100.frx":0D4C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf100.frx":0DBA
         Keys            =   "frmYpmf100.frx":0DD8
         Spin            =   "frmYpmf100.frx":0E22
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnTax_Total 
         Height          =   375
         Left            =   9900
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   180
         Width           =   1000
         _Version        =   65536
         _ExtentX        =   1764
         _ExtentY        =   661
         Calculator      =   "frmYpmf100.frx":0E4A
         Caption         =   "frmYpmf100.frx":0E6A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf100.frx":0ED8
         Keys            =   "frmYpmf100.frx":0EF6
         Spin            =   "frmYpmf100.frx":0F40
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1179653
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnTotal2_Total 
         Height          =   375
         Left            =   8550
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   180
         Width           =   1300
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   661
         Calculator      =   "frmYpmf100.frx":0F68
         Caption         =   "frmYpmf100.frx":0F88
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf100.frx":0FF6
         Keys            =   "frmYpmf100.frx":1014
         Spin            =   "frmYpmf100.frx":105E
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1179653
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnKeep_Total 
         Height          =   375
         Left            =   6690
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   180
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   661
         Calculator      =   "frmYpmf100.frx":1086
         Caption         =   "frmYpmf100.frx":10A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf100.frx":1114
         Keys            =   "frmYpmf100.frx":1132
         Spin            =   "frmYpmf100.frx":117C
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnGtotal_Total 
         Height          =   375
         Left            =   10990
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   180
         Width           =   1400
         _Version        =   65536
         _ExtentX        =   2469
         _ExtentY        =   661
         Calculator      =   "frmYpmf100.frx":11A4
         Caption         =   "frmYpmf100.frx":11C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf100.frx":1232
         Keys            =   "frmYpmf100.frx":1250
         Spin            =   "frmYpmf100.frx":129A
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1179653
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnZandaka_Total 
         Height          =   375
         Left            =   12900
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   180
         Width           =   1300
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   661
         Calculator      =   "frmYpmf100.frx":12C2
         Caption         =   "frmYpmf100.frx":12E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf100.frx":1350
         Keys            =   "frmYpmf100.frx":136E
         Spin            =   "frmYpmf100.frx":13B8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnBcodeCount 
         Height          =   375
         Left            =   1320
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         Calculator      =   "frmYpmf100.frx":13E0
         Caption         =   "frmYpmf100.frx":1400
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf100.frx":146E
         Keys            =   "frmYpmf100.frx":148C
         Spin            =   "frmYpmf100.frx":14D6
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2011365381
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnBrate2_Total 
         Height          =   375
         Left            =   7620
         TabIndex        =   197
         TabStop         =   0   'False
         Top             =   180
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   661
         Calculator      =   "frmYpmf100.frx":14FE
         Caption         =   "frmYpmf100.frx":151E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf100.frx":158C
         Keys            =   "frmYpmf100.frx":15AA
         Spin            =   "frmYpmf100.frx":15F4
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin VB.Label Label1 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "���匏��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   184
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton cmdLogin 
         Caption         =   "�J�ÔN�����ƒS���҂̕ύX"
         Height          =   375
         Left            =   6960
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9960
         TabIndex        =   12
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
         Contents        =   "frmYpmf100.frx":161C
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   10
      Top             =   9600
      Width           =   14835
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   4
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "�ĕ\��(F8)"
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
         rText.left      =   15
         rText.top       =   8
         rText.right     =   97
         rText.bottom    =   27
         Picture         =   "frmYpmf100.frx":1635
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   12600
         TabIndex        =   6
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
         Picture         =   "frmYpmf100.frx":1651
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   10200
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   2355
         _Version        =   262145
         _ExtentX        =   4154
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "�����ꗗ�\(F12)"
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
         rPic.left       =   8
         rPic.top        =   8
         rPic.right      =   16
         rPic.bottom     =   16
         rText.left      =   28
         rText.top       =   8
         rText.right     =   151
         rText.bottom    =   27
         Picture         =   "frmYpmf100.frx":17AB
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   7695
      Left            =   60
      TabIndex        =   9
      Top             =   1320
      Width           =   14835
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   107
         Top             =   6900
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   10
            Left            =   4800
            Picture         =   "frmYpmf100.frx":18BD
            Style           =   1  '���̨���
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   10
            Left            =   4380
            Picture         =   "frmYpmf100.frx":19BF
            Style           =   1  '���̨���
            TabIndex        =   172
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   10
            Left            =   3540
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":1CC9
            Caption         =   "frmYpmf100.frx":1CE9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":1D57
            Keys            =   "frmYpmf100.frx":1D75
            Spin            =   "frmYpmf100.frx":1DBF
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   10
            Left            =   5220
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":1DE7
            Caption         =   "frmYpmf100.frx":1E07
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":1E75
            Keys            =   "frmYpmf100.frx":1E93
            Spin            =   "frmYpmf100.frx":1EDD
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   10
            Left            =   9800
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":1F05
            Caption         =   "frmYpmf100.frx":1F25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":1F93
            Keys            =   "frmYpmf100.frx":1FB1
            Spin            =   "frmYpmf100.frx":1FFB
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   10
            Left            =   8430
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":2023
            Caption         =   "frmYpmf100.frx":2043
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":20B1
            Keys            =   "frmYpmf100.frx":20CF
            Spin            =   "frmYpmf100.frx":2119
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   10
            Left            =   6570
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":2141
            Caption         =   "frmYpmf100.frx":2161
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":21CF
            Keys            =   "frmYpmf100.frx":21ED
            Spin            =   "frmYpmf100.frx":2237
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   10
            Left            =   10870
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":225F
            Caption         =   "frmYpmf100.frx":227F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":22ED
            Keys            =   "frmYpmf100.frx":230B
            Spin            =   "frmYpmf100.frx":2355
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   10
            Left            =   12770
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":237D
            Caption         =   "frmYpmf100.frx":239D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":240B
            Keys            =   "frmYpmf100.frx":2429
            Spin            =   "frmYpmf100.frx":2473
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   10
            Left            =   3960
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":249B
            Caption         =   "frmYpmf100.frx":24BB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":2529
            Keys            =   "frmYpmf100.frx":2547
            Spin            =   "frmYpmf100.frx":2591
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   10
            Left            =   7500
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":25B9
            Caption         =   "frmYpmf100.frx":25D9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":2647
            Keys            =   "frmYpmf100.frx":2665
            Spin            =   "frmYpmf100.frx":26AF
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   10
            Left            =   12320
            TabIndex        =   125
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   10
            Left            =   60
            TabIndex        =   115
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   10
            Left            =   840
            TabIndex        =   114
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   98
         Top             =   6240
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   9
            Left            =   4800
            Picture         =   "frmYpmf100.frx":26D7
            Style           =   1  '���̨���
            TabIndex        =   182
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   9
            Left            =   4380
            Picture         =   "frmYpmf100.frx":27D9
            Style           =   1  '���̨���
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   9
            Left            =   3540
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":2AE3
            Caption         =   "frmYpmf100.frx":2B03
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":2B71
            Keys            =   "frmYpmf100.frx":2B8F
            Spin            =   "frmYpmf100.frx":2BD9
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   9
            Left            =   5220
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":2C01
            Caption         =   "frmYpmf100.frx":2C21
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":2C8F
            Keys            =   "frmYpmf100.frx":2CAD
            Spin            =   "frmYpmf100.frx":2CF7
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   9
            Left            =   9800
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":2D1F
            Caption         =   "frmYpmf100.frx":2D3F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":2DAD
            Keys            =   "frmYpmf100.frx":2DCB
            Spin            =   "frmYpmf100.frx":2E15
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   9
            Left            =   8430
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":2E3D
            Caption         =   "frmYpmf100.frx":2E5D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":2ECB
            Keys            =   "frmYpmf100.frx":2EE9
            Spin            =   "frmYpmf100.frx":2F33
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   9
            Left            =   6570
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":2F5B
            Caption         =   "frmYpmf100.frx":2F7B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":2FE9
            Keys            =   "frmYpmf100.frx":3007
            Spin            =   "frmYpmf100.frx":3051
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   9
            Left            =   10870
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":3079
            Caption         =   "frmYpmf100.frx":3099
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3107
            Keys            =   "frmYpmf100.frx":3125
            Spin            =   "frmYpmf100.frx":316F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   9
            Left            =   12770
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":3197
            Caption         =   "frmYpmf100.frx":31B7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3225
            Keys            =   "frmYpmf100.frx":3243
            Spin            =   "frmYpmf100.frx":328D
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   9
            Left            =   3960
            TabIndex        =   155
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":32B5
            Caption         =   "frmYpmf100.frx":32D5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3343
            Keys            =   "frmYpmf100.frx":3361
            Spin            =   "frmYpmf100.frx":33AB
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   9
            Left            =   7500
            TabIndex        =   195
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":33D3
            Caption         =   "frmYpmf100.frx":33F3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3461
            Keys            =   "frmYpmf100.frx":347F
            Spin            =   "frmYpmf100.frx":34C9
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   12320
            TabIndex        =   124
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   840
            TabIndex        =   106
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   9
            Left            =   60
            TabIndex        =   105
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   89
         Top             =   5580
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   8
            Left            =   4800
            Picture         =   "frmYpmf100.frx":34F1
            Style           =   1  '���̨���
            TabIndex        =   181
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   8
            Left            =   4380
            Picture         =   "frmYpmf100.frx":35F3
            Style           =   1  '���̨���
            TabIndex        =   170
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   8
            Left            =   3540
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":38FD
            Caption         =   "frmYpmf100.frx":391D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":398B
            Keys            =   "frmYpmf100.frx":39A9
            Spin            =   "frmYpmf100.frx":39F3
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   8
            Left            =   5220
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":3A1B
            Caption         =   "frmYpmf100.frx":3A3B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3AA9
            Keys            =   "frmYpmf100.frx":3AC7
            Spin            =   "frmYpmf100.frx":3B11
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   8
            Left            =   9800
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":3B39
            Caption         =   "frmYpmf100.frx":3B59
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3BC7
            Keys            =   "frmYpmf100.frx":3BE5
            Spin            =   "frmYpmf100.frx":3C2F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   8
            Left            =   8430
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":3C57
            Caption         =   "frmYpmf100.frx":3C77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3CE5
            Keys            =   "frmYpmf100.frx":3D03
            Spin            =   "frmYpmf100.frx":3D4D
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   8
            Left            =   6570
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":3D75
            Caption         =   "frmYpmf100.frx":3D95
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3E03
            Keys            =   "frmYpmf100.frx":3E21
            Spin            =   "frmYpmf100.frx":3E6B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   8
            Left            =   10870
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":3E93
            Caption         =   "frmYpmf100.frx":3EB3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":3F21
            Keys            =   "frmYpmf100.frx":3F3F
            Spin            =   "frmYpmf100.frx":3F89
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   8
            Left            =   12770
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":3FB1
            Caption         =   "frmYpmf100.frx":3FD1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":403F
            Keys            =   "frmYpmf100.frx":405D
            Spin            =   "frmYpmf100.frx":40A7
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   8
            Left            =   3960
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":40CF
            Caption         =   "frmYpmf100.frx":40EF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":415D
            Keys            =   "frmYpmf100.frx":417B
            Spin            =   "frmYpmf100.frx":41C5
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   8
            Left            =   7500
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":41ED
            Caption         =   "frmYpmf100.frx":420D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":427B
            Keys            =   "frmYpmf100.frx":4299
            Spin            =   "frmYpmf100.frx":42E3
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   12320
            TabIndex        =   123
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   840
            TabIndex        =   97
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   8
            Left            =   60
            TabIndex        =   96
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   80
         Top             =   4920
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   7
            Left            =   4800
            Picture         =   "frmYpmf100.frx":430B
            Style           =   1  '���̨���
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   7
            Left            =   4380
            Picture         =   "frmYpmf100.frx":440D
            Style           =   1  '���̨���
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   7
            Left            =   3540
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":4717
            Caption         =   "frmYpmf100.frx":4737
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":47A5
            Keys            =   "frmYpmf100.frx":47C3
            Spin            =   "frmYpmf100.frx":480D
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   7
            Left            =   5220
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":4835
            Caption         =   "frmYpmf100.frx":4855
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":48C3
            Keys            =   "frmYpmf100.frx":48E1
            Spin            =   "frmYpmf100.frx":492B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   7
            Left            =   9800
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":4953
            Caption         =   "frmYpmf100.frx":4973
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":49E1
            Keys            =   "frmYpmf100.frx":49FF
            Spin            =   "frmYpmf100.frx":4A49
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   7
            Left            =   8430
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":4A71
            Caption         =   "frmYpmf100.frx":4A91
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":4AFF
            Keys            =   "frmYpmf100.frx":4B1D
            Spin            =   "frmYpmf100.frx":4B67
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   7
            Left            =   6570
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":4B8F
            Caption         =   "frmYpmf100.frx":4BAF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":4C1D
            Keys            =   "frmYpmf100.frx":4C3B
            Spin            =   "frmYpmf100.frx":4C85
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   7
            Left            =   10870
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":4CAD
            Caption         =   "frmYpmf100.frx":4CCD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":4D3B
            Keys            =   "frmYpmf100.frx":4D59
            Spin            =   "frmYpmf100.frx":4DA3
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   7
            Left            =   12770
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":4DCB
            Caption         =   "frmYpmf100.frx":4DEB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":4E59
            Keys            =   "frmYpmf100.frx":4E77
            Spin            =   "frmYpmf100.frx":4EC1
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   7
            Left            =   3960
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":4EE9
            Caption         =   "frmYpmf100.frx":4F09
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":4F77
            Keys            =   "frmYpmf100.frx":4F95
            Spin            =   "frmYpmf100.frx":4FDF
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   7
            Left            =   7500
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":5007
            Caption         =   "frmYpmf100.frx":5027
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":5095
            Keys            =   "frmYpmf100.frx":50B3
            Spin            =   "frmYpmf100.frx":50FD
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   12320
            TabIndex        =   122
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   840
            TabIndex        =   88
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   7
            Left            =   60
            TabIndex        =   87
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   71
         Top             =   4260
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   6
            Left            =   4800
            Picture         =   "frmYpmf100.frx":5125
            Style           =   1  '���̨���
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   6
            Left            =   4380
            Picture         =   "frmYpmf100.frx":5227
            Style           =   1  '���̨���
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   6
            Left            =   3540
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":5531
            Caption         =   "frmYpmf100.frx":5551
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":55BF
            Keys            =   "frmYpmf100.frx":55DD
            Spin            =   "frmYpmf100.frx":5627
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   6
            Left            =   5220
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":564F
            Caption         =   "frmYpmf100.frx":566F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":56DD
            Keys            =   "frmYpmf100.frx":56FB
            Spin            =   "frmYpmf100.frx":5745
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   6
            Left            =   9800
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":576D
            Caption         =   "frmYpmf100.frx":578D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":57FB
            Keys            =   "frmYpmf100.frx":5819
            Spin            =   "frmYpmf100.frx":5863
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   6
            Left            =   8430
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":588B
            Caption         =   "frmYpmf100.frx":58AB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":5919
            Keys            =   "frmYpmf100.frx":5937
            Spin            =   "frmYpmf100.frx":5981
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   6
            Left            =   6570
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":59A9
            Caption         =   "frmYpmf100.frx":59C9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":5A37
            Keys            =   "frmYpmf100.frx":5A55
            Spin            =   "frmYpmf100.frx":5A9F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   6
            Left            =   10870
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":5AC7
            Caption         =   "frmYpmf100.frx":5AE7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":5B55
            Keys            =   "frmYpmf100.frx":5B73
            Spin            =   "frmYpmf100.frx":5BBD
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   6
            Left            =   12770
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":5BE5
            Caption         =   "frmYpmf100.frx":5C05
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":5C73
            Keys            =   "frmYpmf100.frx":5C91
            Spin            =   "frmYpmf100.frx":5CDB
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   6
            Left            =   3960
            TabIndex        =   152
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":5D03
            Caption         =   "frmYpmf100.frx":5D23
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":5D91
            Keys            =   "frmYpmf100.frx":5DAF
            Spin            =   "frmYpmf100.frx":5DF9
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   6
            Left            =   7500
            TabIndex        =   192
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":5E21
            Caption         =   "frmYpmf100.frx":5E41
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":5EAF
            Keys            =   "frmYpmf100.frx":5ECD
            Spin            =   "frmYpmf100.frx":5F17
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   12320
            TabIndex        =   121
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   840
            TabIndex        =   79
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   6
            Left            =   60
            TabIndex        =   78
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Top             =   3600
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   5
            Left            =   4800
            Picture         =   "frmYpmf100.frx":5F3F
            Style           =   1  '���̨���
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   5
            Left            =   4380
            Picture         =   "frmYpmf100.frx":6041
            Style           =   1  '���̨���
            TabIndex        =   167
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   5
            Left            =   3540
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":634B
            Caption         =   "frmYpmf100.frx":636B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":63D9
            Keys            =   "frmYpmf100.frx":63F7
            Spin            =   "frmYpmf100.frx":6441
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   5
            Left            =   5220
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":6469
            Caption         =   "frmYpmf100.frx":6489
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":64F7
            Keys            =   "frmYpmf100.frx":6515
            Spin            =   "frmYpmf100.frx":655F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   5
            Left            =   9800
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":6587
            Caption         =   "frmYpmf100.frx":65A7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":6615
            Keys            =   "frmYpmf100.frx":6633
            Spin            =   "frmYpmf100.frx":667D
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   5
            Left            =   8430
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":66A5
            Caption         =   "frmYpmf100.frx":66C5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":6733
            Keys            =   "frmYpmf100.frx":6751
            Spin            =   "frmYpmf100.frx":679B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   5
            Left            =   6570
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":67C3
            Caption         =   "frmYpmf100.frx":67E3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":6851
            Keys            =   "frmYpmf100.frx":686F
            Spin            =   "frmYpmf100.frx":68B9
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   5
            Left            =   10870
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":68E1
            Caption         =   "frmYpmf100.frx":6901
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":696F
            Keys            =   "frmYpmf100.frx":698D
            Spin            =   "frmYpmf100.frx":69D7
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   5
            Left            =   12770
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":69FF
            Caption         =   "frmYpmf100.frx":6A1F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":6A8D
            Keys            =   "frmYpmf100.frx":6AAB
            Spin            =   "frmYpmf100.frx":6AF5
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   5
            Left            =   3960
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":6B1D
            Caption         =   "frmYpmf100.frx":6B3D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":6BAB
            Keys            =   "frmYpmf100.frx":6BC9
            Spin            =   "frmYpmf100.frx":6C13
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   5
            Left            =   7500
            TabIndex        =   191
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":6C3B
            Caption         =   "frmYpmf100.frx":6C5B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":6CC9
            Keys            =   "frmYpmf100.frx":6CE7
            Spin            =   "frmYpmf100.frx":6D31
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   12320
            TabIndex        =   120
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   5
            Left            =   60
            TabIndex        =   70
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   840
            TabIndex        =   69
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   2940
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   4
            Left            =   4800
            Picture         =   "frmYpmf100.frx":6D59
            Style           =   1  '���̨���
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   4
            Left            =   4380
            Picture         =   "frmYpmf100.frx":6E5B
            Style           =   1  '���̨���
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   4
            Left            =   3540
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":7165
            Caption         =   "frmYpmf100.frx":7185
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":71F3
            Keys            =   "frmYpmf100.frx":7211
            Spin            =   "frmYpmf100.frx":725B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   4
            Left            =   5220
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":7283
            Caption         =   "frmYpmf100.frx":72A3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":7311
            Keys            =   "frmYpmf100.frx":732F
            Spin            =   "frmYpmf100.frx":7379
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   4
            Left            =   9800
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":73A1
            Caption         =   "frmYpmf100.frx":73C1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":742F
            Keys            =   "frmYpmf100.frx":744D
            Spin            =   "frmYpmf100.frx":7497
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   4
            Left            =   8430
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":74BF
            Caption         =   "frmYpmf100.frx":74DF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":754D
            Keys            =   "frmYpmf100.frx":756B
            Spin            =   "frmYpmf100.frx":75B5
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   4
            Left            =   6570
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":75DD
            Caption         =   "frmYpmf100.frx":75FD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":766B
            Keys            =   "frmYpmf100.frx":7689
            Spin            =   "frmYpmf100.frx":76D3
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   4
            Left            =   10870
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":76FB
            Caption         =   "frmYpmf100.frx":771B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":7789
            Keys            =   "frmYpmf100.frx":77A7
            Spin            =   "frmYpmf100.frx":77F1
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   4
            Left            =   12770
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":7819
            Caption         =   "frmYpmf100.frx":7839
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":78A7
            Keys            =   "frmYpmf100.frx":78C5
            Spin            =   "frmYpmf100.frx":790F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   4
            Left            =   3960
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":7937
            Caption         =   "frmYpmf100.frx":7957
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":79C5
            Keys            =   "frmYpmf100.frx":79E3
            Spin            =   "frmYpmf100.frx":7A2D
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   4
            Left            =   7500
            TabIndex        =   190
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":7A55
            Caption         =   "frmYpmf100.frx":7A75
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":7AE3
            Keys            =   "frmYpmf100.frx":7B01
            Spin            =   "frmYpmf100.frx":7B4B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   12320
            TabIndex        =   119
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   4
            Left            =   60
            TabIndex        =   61
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   840
            TabIndex        =   60
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   44
         Top             =   2280
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   3
            Left            =   4800
            Picture         =   "frmYpmf100.frx":7B73
            Style           =   1  '���̨���
            TabIndex        =   176
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   3
            Left            =   4380
            Picture         =   "frmYpmf100.frx":7C75
            Style           =   1  '���̨���
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   3
            Left            =   3540
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":7F7F
            Caption         =   "frmYpmf100.frx":7F9F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":800D
            Keys            =   "frmYpmf100.frx":802B
            Spin            =   "frmYpmf100.frx":8075
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   3
            Left            =   5220
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":809D
            Caption         =   "frmYpmf100.frx":80BD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":812B
            Keys            =   "frmYpmf100.frx":8149
            Spin            =   "frmYpmf100.frx":8193
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   3
            Left            =   9800
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":81BB
            Caption         =   "frmYpmf100.frx":81DB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":8249
            Keys            =   "frmYpmf100.frx":8267
            Spin            =   "frmYpmf100.frx":82B1
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   3
            Left            =   8430
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":82D9
            Caption         =   "frmYpmf100.frx":82F9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":8367
            Keys            =   "frmYpmf100.frx":8385
            Spin            =   "frmYpmf100.frx":83CF
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   3
            Left            =   6570
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":83F7
            Caption         =   "frmYpmf100.frx":8417
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":8485
            Keys            =   "frmYpmf100.frx":84A3
            Spin            =   "frmYpmf100.frx":84ED
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   3
            Left            =   10870
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":8515
            Caption         =   "frmYpmf100.frx":8535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":85A3
            Keys            =   "frmYpmf100.frx":85C1
            Spin            =   "frmYpmf100.frx":860B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   3
            Left            =   12770
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":8633
            Caption         =   "frmYpmf100.frx":8653
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":86C1
            Keys            =   "frmYpmf100.frx":86DF
            Spin            =   "frmYpmf100.frx":8729
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   3
            Left            =   3960
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":8751
            Caption         =   "frmYpmf100.frx":8771
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":87DF
            Keys            =   "frmYpmf100.frx":87FD
            Spin            =   "frmYpmf100.frx":8847
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   3
            Left            =   7500
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":886F
            Caption         =   "frmYpmf100.frx":888F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":88FD
            Keys            =   "frmYpmf100.frx":891B
            Spin            =   "frmYpmf100.frx":8965
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   12320
            TabIndex        =   118
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   3
            Left            =   60
            TabIndex        =   52
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   840
            TabIndex        =   51
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   1620
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   2
            Left            =   4800
            Picture         =   "frmYpmf100.frx":898D
            Style           =   1  '���̨���
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   2
            Left            =   4380
            Picture         =   "frmYpmf100.frx":8A8F
            Style           =   1  '���̨���
            TabIndex        =   164
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   2
            Left            =   3540
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":8D99
            Caption         =   "frmYpmf100.frx":8DB9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":8E27
            Keys            =   "frmYpmf100.frx":8E45
            Spin            =   "frmYpmf100.frx":8E8F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2011365381
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   2
            Left            =   5220
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":8EB7
            Caption         =   "frmYpmf100.frx":8ED7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":8F45
            Keys            =   "frmYpmf100.frx":8F63
            Spin            =   "frmYpmf100.frx":8FAD
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   2
            Left            =   9800
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":8FD5
            Caption         =   "frmYpmf100.frx":8FF5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":9063
            Keys            =   "frmYpmf100.frx":9081
            Spin            =   "frmYpmf100.frx":90CB
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   2
            Left            =   8430
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":90F3
            Caption         =   "frmYpmf100.frx":9113
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":9181
            Keys            =   "frmYpmf100.frx":919F
            Spin            =   "frmYpmf100.frx":91E9
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   2
            Left            =   6570
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":9211
            Caption         =   "frmYpmf100.frx":9231
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":929F
            Keys            =   "frmYpmf100.frx":92BD
            Spin            =   "frmYpmf100.frx":9307
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   2
            Left            =   10870
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":932F
            Caption         =   "frmYpmf100.frx":934F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":93BD
            Keys            =   "frmYpmf100.frx":93DB
            Spin            =   "frmYpmf100.frx":9425
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   2
            Left            =   12770
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":944D
            Caption         =   "frmYpmf100.frx":946D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":94DB
            Keys            =   "frmYpmf100.frx":94F9
            Spin            =   "frmYpmf100.frx":9543
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":956B
            Caption         =   "frmYpmf100.frx":958B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":95F9
            Keys            =   "frmYpmf100.frx":9617
            Spin            =   "frmYpmf100.frx":9661
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2011365381
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   2
            Left            =   7500
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":9689
            Caption         =   "frmYpmf100.frx":96A9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":9717
            Keys            =   "frmYpmf100.frx":9735
            Spin            =   "frmYpmf100.frx":977F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   12320
            TabIndex        =   117
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Index           =   2
            Left            =   60
            TabIndex        =   43
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   840
            TabIndex        =   42
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6495
         Left            =   14400
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1080
         Width           =   315
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   14175
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   1
            Left            =   4800
            Picture         =   "frmYpmf100.frx":97A7
            Style           =   1  '���̨���
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   1
            Left            =   4380
            Picture         =   "frmYpmf100.frx":98A9
            Style           =   1  '���̨���
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   1
            Left            =   3540
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":9BB3
            Caption         =   "frmYpmf100.frx":9BD3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":9C41
            Keys            =   "frmYpmf100.frx":9C5F
            Spin            =   "frmYpmf100.frx":9CA9
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2011365381
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   1
            Left            =   5220
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":9CD1
            Caption         =   "frmYpmf100.frx":9CF1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":9D5F
            Keys            =   "frmYpmf100.frx":9D7D
            Spin            =   "frmYpmf100.frx":9DC7
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTax 
            Height          =   375
            Index           =   1
            Left            =   9800
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":9DEF
            Caption         =   "frmYpmf100.frx":9E0F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":9E7D
            Keys            =   "frmYpmf100.frx":9E9B
            Spin            =   "frmYpmf100.frx":9EE5
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal2 
            Height          =   375
            Index           =   1
            Left            =   8430
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":9F0D
            Caption         =   "frmYpmf100.frx":9F2D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":9F9B
            Keys            =   "frmYpmf100.frx":9FB9
            Spin            =   "frmYpmf100.frx":A003
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeep 
            Height          =   375
            Index           =   1
            Left            =   6570
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":A02B
            Caption         =   "frmYpmf100.frx":A04B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":A0B9
            Keys            =   "frmYpmf100.frx":A0D7
            Spin            =   "frmYpmf100.frx":A121
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   1
            Left            =   10870
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   180
            Width           =   1400
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":A149
            Caption         =   "frmYpmf100.frx":A169
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":A1D7
            Keys            =   "frmYpmf100.frx":A1F5
            Spin            =   "frmYpmf100.frx":A23F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   1
            Left            =   12770
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   180
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":A267
            Caption         =   "frmYpmf100.frx":A287
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":A2F5
            Keys            =   "frmYpmf100.frx":A313
            Spin            =   "frmYpmf100.frx":A35D
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   1
            Left            =   3960
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":A385
            Caption         =   "frmYpmf100.frx":A3A5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":A413
            Keys            =   "frmYpmf100.frx":A431
            Spin            =   "frmYpmf100.frx":A47B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999
            MinValue        =   -999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2011365381
            Value           =   999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   1
            Left            =   7500
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf100.frx":A4A3
            Caption         =   "frmYpmf100.frx":A4C3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf100.frx":A531
            Keys            =   "frmYpmf100.frx":A54F
            Spin            =   "frmYpmf100.frx":A599
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   12320
            TabIndex        =   116
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblBname 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   840
            TabIndex        =   20
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BorderStyle     =   1  '����
            Caption         =   "9999"
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
            Left            =   60
            TabIndex        =   19
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Width           =   14175
         Begin VB.Label Label1 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�萔��"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Index           =   10
            Left            =   7620
            TabIndex        =   187
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�e"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   3960
            TabIndex        =   147
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label1 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�O�񖘎c��"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   12770
            TabIndex        =   134
            Top             =   240
            Width           =   1300
         End
         Begin VB.Label Label1 
            Alignment       =   2  '��������
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Index           =   1
            Left            =   3540
            TabIndex        =   132
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label1 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�����v"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   10870
            TabIndex        =   33
            Top             =   240
            Width           =   1400
         End
         Begin VB.Label Label1 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�ێ��Ǘ���"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Index           =   6
            Left            =   6570
            TabIndex        =   32
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�v"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   8520
            TabIndex        =   31
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�����"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   9720
            TabIndex        =   30
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  '�E����
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�������v"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   5220
            TabIndex        =   29
            Top             =   240
            Width           =   1300
         End
         Begin VB.Label Label1 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "���@��"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1425
         End
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   15120
      TabIndex        =   0
      Top             =   0
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf100.frx":A5C1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf100.frx":A62F
      Key             =   "frmYpmf100.frx":A64D
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
      TabIndex        =   7
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf100.frx":A691
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf100.frx":A6FF
      Key             =   "frmYpmf100.frx":A71D
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
      TabIndex        =   8
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf100.frx":A761
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf100.frx":A7CF
      Key             =   "frmYpmf100.frx":A7ED
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
   Begin CSCmdLibCtl.CSCmdBtn cmdRelease 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   12780
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   120
      Width           =   2115
      _Version        =   262145
      _ExtentX        =   3731
      _ExtentY        =   767
      _StockProps     =   15
      Caption         =   "���吸�Z�S����"
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
      rText.left      =   12
      rText.top       =   6
      rText.right     =   128
      rText.bottom    =   25
      Picture         =   "frmYpmf100.frx":A831
   End
End
Attribute VB_Name = "frmYpmf100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DETAIL_MAX = 10               '���ו\���s��
Private Const BACK_COLOR_ON = &HFF0000
Private Const BACK_COLOR_OFF = &H8000000F
Private Const BACK_COLOR_CAUTION = &HFF&
Private Const DIV_NAME1 = "��"
Private Const DIV_NAME2 = "��"
Private Const DIV_NAME3 = "����"

Private Type Detail_Record
    Bcode As String
    Bname As String
    Num As Integer
    F As Long
    Total As Currency
    Tax As Currency
    Total2 As Currency
    Keep As Currency
    Brate2 As Currency '201107
    Gtotal As Currency
    Div As String
    Zandaka As Currency
End Type
Private m_typDetail_Rec() As Detail_Record

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

Private Sub cboBcode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Call Form_KeyDown(KeyCode, Shift)

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

'�ځ@�I�@�@�F
'���@���@�@�F��ʃN���A�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Sub cmdClear_Click()

    Call FieldsClear(1)

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F����N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

'    If MsgBox("������܂����H", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    '����p���[�N�쐬
    If MakePrintWork() = False Then Exit Sub
    '����v���r���[
    If ActiveReportPrint(0) = False Then Exit Sub
    
    Exit Sub
    
cmdExecute_Click_Err:

    Call MsgBox("����N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�I���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdF_Search_Click(Index As Integer)

    frmView.m_intBode = lblBcode(Index).Caption
    frmView.m_intBnum = imnNum(Index).Value
    frmView.m_strBname = lblBname(Index).Caption
    frmView.Show vbModal

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
        Call FieldsClear(1)
    End If
    Unload frmLogin
    
    Exit Sub

cmdLogin_Click_Err:

    Call MsgBox("�J�ÔN�����ƒS���҂̕ύX�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdLogin_Click_Err")

End Sub

Private Sub cmdPrint_Click(Index As Integer)

    If MsgBox("���吸�Z�`�[��������܂����H", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    Dim strArg As String
    strArg = lblOdate.Caption & "," & g_strPcode & "," & g_strPname & "," & lblBcode(Index).Caption & "," & imnNum(Index).Value
'    Call Shell(g_clsReg.Bin & "\YPMF050.exe " & strArg, vbNormalFocus)
    Call Shell(g_clsReg.Bin & "\YPMF050.exe " & strArg, vbMaximizedFocus)
    

End Sub

Private Sub cmdRelease_Click()

    If MsgBox("���吸�Z��S�������܂����H", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub
    DoEvents
    If MsgBox("�{���ɂ�낵���ł����H", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub
    DoEvents
    
    If Release_Data() = False Then Exit Sub
    Call FieldsClear(1)

End Sub

Private Sub cmdSearch_Click()

    Dim intIndex1 As Integer

    On Error GoTo cmdSearch_Click_Err
    
'    If Trim(cboBcode(0).Text) = "" Then Exit Sub
'    If Trim(cboBcode(1).Text) = "" Then Exit Sub
    
    For intIndex1 = 1 To DETAIL_MAX
        Call Detail_Clear(intIndex1)
    Next intIndex1
    
    Erase m_typDetail_Rec   '�z�񏉊���
    
    imnTotal_Total.Value = 0
    imnTax_Total.Value = 0
    imnTotal2_Total.Value = 0
    imnKeep_Total.Value = 0
    imnBrate2_Total.Value = 0   '201107
    imnGtotal_Total.Value = 0
    imnZandaka_Total.Value = 0
    
    If Trim(cboBcode(0).Text) = "" And Trim(cboBcode(1).Text) = "" Then
        Call Detail_SetData(0)
    Else
        Call Detail_SetData(1)
    End If
    Call Detail_Dislplay(1)
    Call Detail_ScrollBar
    
    If UBound(m_typDetail_Rec) <= 0 Then
        Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "")
    End If
    
    Exit Sub
    
cmdSearch_Click_Err:

    Call MsgBox("��ʃN���A�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdSearch_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
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
        Case vbKeyF2
        Case vbKeyHome
        Case vbKeyPageUp
            If VScroll1.Value > 1 Then
                VScroll1.Value = VScroll1 - 1
            End If
        Case vbKeyPageDown
            If (VScroll1.Value + 1) <= VScroll1.Max Then
                VScroll1.Value = VScroll1 + 1
            End If
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
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

'    Me.Caption = SYSTEM_NAME & "-" & "���吸�Z��"
    Me.Caption = "���吸�Z��"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    Call FieldsClear(0)
    Timer1.Enabled = True
    
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
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
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
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Sub FieldsClear(intKubun As Integer)

    Dim intIndex1 As Integer

    On Error GoTo FieldsClear_Err
    
    For intIndex1 = 1 To DETAIL_MAX
        Call Detail_Clear(intIndex1)
    Next intIndex1
    
    Erase m_typDetail_Rec   '�z�񏉊���
    
    cboBcode(0).Text = ""
    cboBcode(1).Text = ""
    lblScode_Name(0).Caption = ""
    lblScode_Name(1).Caption = ""
    
    imnBcodeCount.Value = 0
    imnTotal_Total.Value = 0
    imnTax_Total.Value = 0
    imnTotal2_Total.Value = 0
    imnKeep_Total.Value = 0
    imnBrate2_Total.Value = 0   '201107
    imnGtotal_Total.Value = 0
    imnZandaka_Total.Value = 0
    
    If intKubun = 1 Then
        Call Detail_SetData(0)
        Call Detail_Dislplay(1)
        Call Detail_ScrollBar
        
        If UBound(m_typDetail_Rec) <= 0 Then
            Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "")
        End If
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
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
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

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cboBcode(0).SetFocus

End Sub

'�ځ@�I�@�@�F���׃��[�N�փf�[�^�Z�b�g
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�FintFlg 0:�S�f�[�^ 1:����w��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Function Detail_SetData(intFlg As Integer) As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim adoRecordset3 As New ADODB.Recordset
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intRecordCount As Integer
    Dim blnFlg As Boolean
    Dim typDetail_Sort() As Detail_Record
    Dim strBuff As String

    On Error GoTo Detail_SetData_Err
    
    Detail_SetData = False
    
    Screen.MousePointer = vbHourglass
    
    '������
    ReDim typDetail_Sort(0): ReDim m_typDetail_Rec(0)
    
'********** �������׃f�[�^ **********
    
    If intFlg = 0 Then
        strSQL = "{call sp_YPMF1001;1('" & Global_Get_NumericDay(lblOdate.Caption) & "')}"
    Else
        strSQL = "{call sp_YPMF1001;2('" & Global_Get_NumericDay(lblOdate.Caption) & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    End If
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        '���吸�Z�f�[�^
        strSQL = "SELECT * FROM DT041" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                 " AND Num = " & adoRecordset1.Fields("Bnum") & _
                 " ORDER BY Odate,Bcode,Num"
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoRecordset2.EOF = False Then
            Do While Not adoRecordset2.EOF
                intIndex1 = UBound(typDetail_Sort) + 1
                ReDim Preserve typDetail_Sort(intIndex1)
                typDetail_Sort(intIndex1).Bcode = IIf(IsNull(adoRecordset2.Fields("Bcode")), "", adoRecordset2.Fields("Bcode"))
                typDetail_Sort(intIndex1).Bname = IIf(IsNull(adoRecordset2.Fields("Bname")), "", adoRecordset2.Fields("Bname"))
                typDetail_Sort(intIndex1).Num = adoRecordset2.Fields("Num")
                typDetail_Sort(intIndex1).Total = typDetail_Sort(intIndex1).Total + CCur(adoRecordset2.Fields("Total"))
                typDetail_Sort(intIndex1).Tax = typDetail_Sort(intIndex1).Tax + CCur(adoRecordset2.Fields("Tax"))
                '202308 �C���{�C�X�Ή�
                'typDetail_Sort(intIndex1).Total2 = typDetail_Sort(intIndex1).Total + typDetail_Sort(intIndex1).Tax
                typDetail_Sort(intIndex1).Keep = typDetail_Sort(intIndex1).Keep + CCur(adoRecordset2.Fields("Keep"))
                typDetail_Sort(intIndex1).Gtotal = typDetail_Sort(intIndex1).Gtotal + CCur(adoRecordset2.Fields("Gtotal"))
                '�����敪
                If Not IsNull(adoRecordset2.Fields("Rdiv")) Then
                    If adoRecordset2.Fields("Rdiv") = PAYMENT_OFF Then
                        typDetail_Sort(intIndex1).Div = DIV_NAME3
                    ElseIf adoRecordset2.Fields("Rdiv") = PAYMENT_ON Then
                        typDetail_Sort(intIndex1).Div = DIV_NAME2
                    End If
                Else
                    typDetail_Sort(intIndex1).Div = DIV_NAME1
                End If
                
                '201107
                typDetail_Sort(intIndex1).Brate2 = typDetail_Sort(intIndex1).Brate2 + CCur(adoRecordset2.Fields("Brate2"))
                '202308 �C���{�C�X�Ή�
                typDetail_Sort(intIndex1).Total2 = typDetail_Sort(intIndex1).Total + typDetail_Sort(intIndex1).Keep + typDetail_Sort(intIndex1).Brate2
                adoRecordset2.MoveNext
            Loop
        Else
            intIndex1 = UBound(typDetail_Sort) + 1
            ReDim Preserve typDetail_Sort(intIndex1)
            typDetail_Sort(intIndex1).Bcode = IIf(IsNull(adoRecordset1.Fields("Bcode")), "", adoRecordset1.Fields("Bcode"))
            typDetail_Sort(intIndex1).Bname = Global_Get_Bname(g_clsAdoSQL, typDetail_Sort(intIndex1).Bcode, lblOdate.Caption, strBuff)
            typDetail_Sort(intIndex1).Num = 0
            typDetail_Sort(intIndex1).Total = 0
            typDetail_Sort(intIndex1).Tax = 0
            typDetail_Sort(intIndex1).Total2 = 0
            typDetail_Sort(intIndex1).Keep = 0
            typDetail_Sort(intIndex1).Gtotal = 0
            typDetail_Sort(intIndex1).Div = DIV_NAME1
            '201107
            typDetail_Sort(intIndex1).Brate2 = 0
        End If
        adoRecordset2.Close
        
        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close
    
'********** �������׃f�[�^ **********

    If intFlg = 0 Then
        strSQL = "{call sp_YPMF1002;1('" & lblOdate.Caption & "')}"
    Else
        strSQL = "{call sp_YPMF1002;2('" & lblOdate.Caption & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    End If
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        '���׃f�[�^�ɑ��݂��邩�`�F�b�N
        blnFlg = True
        For intIndex1 = 1 To UBound(typDetail_Sort)
            If CInt(typDetail_Sort(intIndex1).Bcode) = CInt(adoRecordset1.Fields("Bcode")) And _
               CInt(typDetail_Sort(intIndex1).Num) = CInt(adoRecordset1.Fields("Bnum")) Then
                blnFlg = False
                Exit For
            End If
        Next intIndex1

        If blnFlg = True Then
            '���吸�Z�f�[�^
            strSQL = "SELECT * FROM DT041" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                     " AND Num = " & adoRecordset1.Fields("Bnum") & _
                     " ORDER BY Odate,Bcode,Num"
            adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoRecordset2.EOF = False Then
                Do While Not adoRecordset2.EOF
                    intIndex1 = UBound(typDetail_Sort) + 1
                    ReDim Preserve typDetail_Sort(intIndex1)
                    typDetail_Sort(intIndex1).Bcode = IIf(IsNull(adoRecordset2.Fields("Bcode")), "", adoRecordset2.Fields("Bcode"))
                    typDetail_Sort(intIndex1).Bname = IIf(IsNull(adoRecordset2.Fields("Bname")), "", adoRecordset2.Fields("Bname"))
                    typDetail_Sort(intIndex1).Num = adoRecordset2.Fields("Num")
                    typDetail_Sort(intIndex1).Total = typDetail_Sort(intIndex1).Total + CCur(adoRecordset2.Fields("Total"))
                    typDetail_Sort(intIndex1).Tax = typDetail_Sort(intIndex1).Tax + CCur(adoRecordset2.Fields("Tax"))
                    typDetail_Sort(intIndex1).Total2 = typDetail_Sort(intIndex1).Total + typDetail_Sort(intIndex1).Tax
                    typDetail_Sort(intIndex1).Keep = typDetail_Sort(intIndex1).Keep + CCur(adoRecordset2.Fields("Keep"))
                    typDetail_Sort(intIndex1).Gtotal = typDetail_Sort(intIndex1).Gtotal + CCur(adoRecordset2.Fields("Gtotal"))
                    typDetail_Sort(intIndex1).Div = DIV_NAME2
                    '201107
                    typDetail_Sort(intIndex1).Brate2 = typDetail_Sort(intIndex1).Brate2 + CCur(adoRecordset2.Fields("Brate2"))
                    
                    adoRecordset2.MoveNext
                Loop
            Else
                intIndex1 = UBound(typDetail_Sort) + 1
                ReDim Preserve typDetail_Sort(intIndex1)
                typDetail_Sort(intIndex1).Bcode = IIf(IsNull(adoRecordset1.Fields("Bcode")), "", adoRecordset1.Fields("Bcode"))
                typDetail_Sort(intIndex1).Bname = Global_Get_Bname(g_clsAdoSQL, typDetail_Sort(intIndex1).Bcode, lblOdate.Caption, strBuff)
                typDetail_Sort(intIndex1).Num = 0
                typDetail_Sort(intIndex1).Total = 0
                typDetail_Sort(intIndex1).Tax = 0
                typDetail_Sort(intIndex1).Total2 = 0
                typDetail_Sort(intIndex1).Keep = 0
                typDetail_Sort(intIndex1).Gtotal = 0
                typDetail_Sort(intIndex1).Div = DIV_NAME1
                '201107
                typDetail_Sort(intIndex1).Brate2 = 0
            End If
            adoRecordset2.Close
        End If

        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close
    
'********** ���吸�Z�f�[�^ **********
    
    If intFlg = 0 Then
        strSQL = "{call sp_YPMF1003;1('" & lblOdate.Caption & "')}"
    Else
        strSQL = "{call sp_YPMF1003;2('" & lblOdate.Caption & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    End If
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        '���׃f�[�^�ɑ��݂��邩�`�F�b�N
        blnFlg = True
        For intIndex1 = 1 To UBound(typDetail_Sort)
            If CInt(typDetail_Sort(intIndex1).Bcode) = CInt(adoRecordset1.Fields("Bcode")) Then
                blnFlg = False
                Exit For
            End If
        Next intIndex1
        
        If blnFlg = True And chkFlg.Value = 1 Then
            '����͐��Z�f�[�^���Ȃ��ꍇ�ɁA�O��܂ł̎c��������ꍇ�ɒǉ�����
            intIndex1 = UBound(typDetail_Sort) + 1
            ReDim Preserve typDetail_Sort(intIndex1)
            typDetail_Sort(intIndex1).Bcode = IIf(IsNull(adoRecordset1.Fields("Bcode")), "", adoRecordset1.Fields("Bcode"))
            typDetail_Sort(intIndex1).Bname = Global_Get_Bname(g_clsAdoSQL, typDetail_Sort(intIndex1).Bcode, lblOdate.Caption, strBuff)
            typDetail_Sort(intIndex1).Num = 0
            typDetail_Sort(intIndex1).Total = 0
            typDetail_Sort(intIndex1).Tax = 0
            typDetail_Sort(intIndex1).Total2 = 0
            typDetail_Sort(intIndex1).Keep = 0
            typDetail_Sort(intIndex1).Gtotal = 0
            typDetail_Sort(intIndex1).Div = DIV_NAME3
            typDetail_Sort(intIndex1).Zandaka = IIf(IsNull(adoRecordset1.Fields("Zandaka")), 0, adoRecordset1.Fields("Zandaka"))
            '201107
            typDetail_Sort(intIndex1).Brate2 = 0
            
            '�����f�[�^�I�[�v��
            strSQL = "{call sp_YPMF1004;1('" & lblOdate.Caption & "'," & typDetail_Sort(intIndex1).Bcode & ")}"
            adoRecordset3.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoRecordset3.EOF = False Then
                If Not IsNull(adoRecordset3.Fields("Ptotal")) Then
                    typDetail_Sort(intIndex1).Zandaka = typDetail_Sort(intIndex1).Zandaka - CCur(adoRecordset3.Fields("Ptotal"))
                End If
            End If
            adoRecordset3.Close
        ElseIf blnFlg = False Then
            '�X�V
            typDetail_Sort(intIndex1).Zandaka = IIf(IsNull(adoRecordset1.Fields("Zandaka")), 0, adoRecordset1.Fields("Zandaka"))
            typDetail_Sort(intIndex1).Div = DIV_NAME3
            
            '�����f�[�^�I�[�v��
            strSQL = "{call sp_YPMF1004;1('" & lblOdate.Caption & "'," & typDetail_Sort(intIndex1).Bcode & ")}"
            adoRecordset3.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoRecordset3.EOF = False Then
                If Not IsNull(adoRecordset3.Fields("Ptotal")) Then
                    typDetail_Sort(intIndex1).Zandaka = typDetail_Sort(intIndex1).Zandaka - CCur(adoRecordset3.Fields("Ptotal"))
                End If
            End If
            adoRecordset3.Close
        End If
        
        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
    '����R�[�h�A�񐔂Ń\�[�g
    Call Detail_Sort(typDetail_Sort, m_typDetail_Rec)
    
    'F�����̎擾
    Call Get_F
    
    '���v�Ȃǂ̌v�Z
    Call Calc_Total
    
    Screen.MousePointer = vbDefault
    
    Detail_SetData = True
    
    Exit Function

Detail_SetData_Err:

    Detail_SetData = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�t�B�[���h�Z�b�g�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_SetData_Err")

End Function

'�ځ@�I�@�@�F���v�̌v�Z
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Sub Calc_Total()

    Dim curBuff(7) As Currency
    Dim intIndex1 As Integer

    On Error GoTo Calc_Total_Err
    
    curBuff(0) = 0: curBuff(1) = 0: curBuff(2) = 0: curBuff(3) = 0: curBuff(4) = 0: curBuff(5) = 0: curBuff(6) = 0: curBuff(7) = 0
    
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
        If m_typDetail_Rec(intIndex1).Num = 1 Then
            curBuff(0) = curBuff(0) + 1
        End If
    
        curBuff(1) = curBuff(1) + m_typDetail_Rec(intIndex1).Total
        curBuff(2) = curBuff(2) + m_typDetail_Rec(intIndex1).Tax
        curBuff(3) = curBuff(3) + m_typDetail_Rec(intIndex1).Total2
        curBuff(4) = curBuff(4) + m_typDetail_Rec(intIndex1).Keep
        curBuff(5) = curBuff(5) + m_typDetail_Rec(intIndex1).Brate2
        curBuff(6) = curBuff(6) + m_typDetail_Rec(intIndex1).Gtotal
        curBuff(7) = curBuff(7) + m_typDetail_Rec(intIndex1).Zandaka
        
    Next intIndex1
    
    imnBcodeCount.Value = curBuff(0)
    imnTotal_Total.Value = curBuff(1)
    imnTax_Total.Value = curBuff(2)
    imnTotal2_Total.Value = curBuff(3)
    imnKeep_Total.Value = curBuff(4)
    imnBrate2_Total.Value = curBuff(5)
    imnGtotal_Total.Value = curBuff(6)
    imnZandaka_Total.Value = curBuff(7)
    
    Exit Sub
    
Calc_Total_Err:

    Call MsgBox("���v�̌v�Z�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Total_Err")

End Sub

'�ځ@�I�@�@�F���ׂ̕\��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F(intStartLine�F�\���J�n�s�ԍ�)
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Function Detail_Dislplay(intStartLine As Integer) As Boolean

    Dim intIndex1 As Integer
    Dim intPostion As Integer

    On Error GoTo Detail_Dislplay_Err
    
    Detail_Dislplay = False
    
    Screen.MousePointer = vbHourglass
    
    intPostion = intStartLine
    For intIndex1 = 1 To DETAIL_MAX
        '���ׂ̂P�s�N���A
        Call Detail_Clear(intIndex1)
        
        If intPostion <= UBound(m_typDetail_Rec) Then
            fraDetail(intIndex1).Visible = True
        
            lblBcode(intIndex1).Caption = m_typDetail_Rec(intPostion).Bcode
            lblBname(intIndex1).Caption = m_typDetail_Rec(intPostion).Bname
            imnNum(intIndex1).Value = m_typDetail_Rec(intPostion).Num
            imnF(intIndex1).Value = m_typDetail_Rec(intPostion).F
            imnTotal(intIndex1).Value = m_typDetail_Rec(intPostion).Total
            imnTax(intIndex1).Value = m_typDetail_Rec(intPostion).Tax
            imnTotal2(intIndex1).Value = m_typDetail_Rec(intPostion).Total2
            imnKeep(intIndex1).Value = m_typDetail_Rec(intPostion).Keep
            imnGtotal(intIndex1).Value = m_typDetail_Rec(intPostion).Gtotal
            lblDiv(intIndex1).Caption = m_typDetail_Rec(intPostion).Div
            If lblDiv(intIndex1).Caption = DIV_NAME1 Then
                fraDetail(intIndex1).BackColor = BACK_COLOR_OFF
            ElseIf lblDiv(intIndex1).Caption = DIV_NAME2 Then
                fraDetail(intIndex1).BackColor = BACK_COLOR_ON
            ElseIf lblDiv(intIndex1).Caption = DIV_NAME3 Then
                fraDetail(intIndex1).BackColor = BACK_COLOR_CAUTION
            End If
            imnZandaka(intIndex1).Value = m_typDetail_Rec(intPostion).Zandaka
            '201107
            imnBrate2(intIndex1).Value = m_typDetail_Rec(intPostion).Brate2
        End If
        intPostion = intPostion + 1
    Next intIndex1
    
    Screen.MousePointer = vbDefault
    
    Detail_Dislplay = True
    
    Exit Function
    
Detail_Dislplay_Err:

    Screen.MousePointer = vbDefault
    Detail_Dislplay = False
    Call MsgBox("���ׂ̕\���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Dislplay_Err")

End Function

'�ځ@�I�@�@�F���ׂ̂P�s�N���A
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F(intClearLine:��Έʒu)
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Function Detail_Clear(intClearLine As Integer) As Boolean

    On Error GoTo Detail_Clear_Err
    
    lblBcode(intClearLine).Caption = ""
    lblBname(intClearLine).Caption = ""
    imnNum(intClearLine).Value = 0
    imnF(intClearLine).Value = 0
    imnTotal(intClearLine).Value = 0
    imnTax(intClearLine).Value = 0
    imnTotal2(intClearLine).Value = 0
    imnKeep(intClearLine).Value = 0
    imnGtotal(intClearLine).Value = 0
    lblDiv(intClearLine).Caption = ""
    imnZandaka(intClearLine).Value = 0
    
    '201107
    imnBrate2(intClearLine).Value = 0
    
    fraDetail(intClearLine).BackColor = BACK_COLOR_OFF
    fraDetail(intClearLine).Visible = False
    
    Detail_Clear = True
    
    Exit Function
    
Detail_Clear_Err:

    Detail_Clear = False
    Call MsgBox("���ׂ̂P�s�N���A�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Clear_Err")

End Function

'�ځ@�I�@�@�F�X�N���[���o�[�̐���
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Function Detail_ScrollBar() As Boolean

    Dim intMax As Integer

    On Error GoTo Detail_ScrollBar_Err
    
    Detail_ScrollBar = False
        
    VScroll1.Tag = "EventFalse"
    If UBound(m_typDetail_Rec) > 0 Then
        VScroll1.Max = UBound(m_typDetail_Rec)
    Else
        VScroll1.Max = 1
    End If
    VScroll1.Min = 1
    VScroll1.Value = 1
    VScroll1.Tag = ""

    Detail_ScrollBar = True
    
    Exit Function
    
Detail_ScrollBar_Err:

    Detail_ScrollBar = False
    Call MsgBox("�X�N���[���o�[�̐���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_ScrollBar_Err")

End Function

Private Sub lblDiv_DblClick(Index As Integer)

    On Error GoTo lblDiv_DblClick_Err

'    If Trim(lblDiv(Index).Caption) <> DIV_NAME2 Then Exit Sub
    
    If MsgBox("����R�[�h�F" & lblBcode(Index).Caption & "�̐��Z���������܂����H", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub
    
    If IsNumeric(lblBcode(Index).Caption) Then
        If Release_Data(CInt(lblBcode(Index).Caption)) = False Then Exit Sub
    End If
    Call FieldsClear(1)

    Exit Sub

lblDiv_DblClick_Err:

    Call MsgBox("���Z�����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "lblDiv_DblClick_Err")
                
End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False
    Call FieldsClear(1)

End Sub

Private Sub VScroll1_Change()

    If VScroll1.Tag = "EventFalse" Then Exit Sub

    Call Detail_Dislplay(VScroll1.Value)

End Sub

'�ځ@�I�@�@�F���ׂ̃\�[�g
'���@���@�@�F����R�[�h�A�񐔂̏���
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Function Detail_Sort(ByRef Before() As Detail_Record, ByRef After() As Detail_Record) As Boolean

    Dim intIndex1 As Integer
    Dim intPostion As Integer
    Dim work(1) As Detail_Record     '���[�N

    On Error GoTo Detail_Sort_Err
    
    If UBound(Before) <= 0 Then Exit Function
    
    '�o�u���\�[�g
    For intIndex1 = UBound(Before) To 1 Step -1
        For intPostion = 1 To intIndex1 - 1
            If CInt(Before(intPostion).Bcode) >= CInt(Before(intPostion + 1).Bcode) And CInt(Before(intPostion).Num) >= CInt(Before(intPostion + 1).Num) Then
                work(1) = Before(intPostion)
                Before(intPostion) = Before(intPostion + 1)
                Before(intPostion + 1) = work(1)
            End If
        Next intPostion
    Next intIndex1
    
    '�z��R�s�[
    ReDim After(UBound(Before))
    After = Before
    
    Detail_Sort = True
    
    Exit Function
    
Detail_Sort_Err:

    Detail_Sort = False
    Call MsgBox("���ׂ̃\�[�g�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Sort_Err")

End Function

'�ځ@�I�@�@�F�o�͍ς݂̉���
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F�Q�O�O�R�^�O�Q�^�Q�P�@�����f�[�^�̉���
'
Private Function Release_Data(Optional intBcode As Variant) As Boolean

    Dim strSQL As String

    On Error GoTo Release_Data_Err
    
    Screen.MousePointer = vbHourglass
    
    Release_Data = False
    
    With g_clsAdoSQL.Connection
        .BeginTrans
        
        If IsMissing(intBcode) = False Then
            '********** �w�蔃��̉��� **********
        
            '�������׃f�[�^
            strSQL = "UPDATE DT021" & _
                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                     " Bnum = 0" & _
                     " WHERE LEFT(Ocode,8) = '" & Global_Get_NumericDay(lblOdate.Caption) & "'" & _
                     " AND Bcode = " & intBcode
            .Execute strSQL
        
            '��t���׃f�[�^
            strSQL = "UPDATE DT011" & _
                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                     " Bnum = 0" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Bcode = " & intBcode
            .Execute strSQL
            
            '�����f�[�^
            strSQL = "UPDATE DT031" & _
                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                     " Bnum = 0" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Bcode = " & intBcode
            .Execute strSQL
        
            '���吸�Z�f�[�^�̍폜
            strSQL = "DELETE FROM DT041" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Bcode = " & intBcode
            .Execute strSQL
        Else
            '********** �S������ **********
        
            '�������׃f�[�^
            strSQL = "UPDATE DT021" & _
                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                     " Bnum = 0" & _
                     " WHERE LEFT(Ocode,8) = '" & Global_Get_NumericDay(lblOdate.Caption) & "'"
            .Execute strSQL
        
            '��t���׃f�[�^
            strSQL = "UPDATE DT011" & _
                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                     " Bnum = 0" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'"
            .Execute strSQL
        
            '�����f�[�^
            strSQL = "UPDATE DT031" & _
                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                     " Bnum = 0" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'"
            .Execute strSQL
        
            '���吸�Z�f�[�^�̍폜
            strSQL = "DELETE FROM DT041" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'"
            .Execute strSQL
        End If
    
        .CommitTrans
    End With
    
    Screen.MousePointer = vbDefault
    
    Release_Data = True
    
    Exit Function
    
Release_Data_Err:

    Release_Data = False
    g_clsAdoSQL.Connection.RollbackTrans
    Call MsgBox("�����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Release_Data_Err")

End Function

'�ځ@�I�@�@�FF�����̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Sub Get_F()

    Dim intIndex1 As Integer
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strOdate As String

    On Error GoTo Get_F_Err
    
    strOdate = CStr(Global_Get_NumericDay(lblOdate.Caption))
    
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
        If m_typDetail_Rec(intIndex1).Num = 0 Then
            '�������׃f�[�^
            strSQL = "SELECT * FROM DT021" & _
                     " WHERE LEFT(Ocode,8) = '" & strOdate & "'" & _
                     " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode & _
                     " AND (Bnum IS NULL OR Bnum = " & m_typDetail_Rec(intIndex1).Num & ")"
        Else
            '�������׃f�[�^
            strSQL = "SELECT * FROM DT021" & _
                     " WHERE LEFT(Ocode,8) = '" & strOdate & "'" & _
                     " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode & _
                     " AND Bnum = " & m_typDetail_Rec(intIndex1).Num
        End If
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoRecordset1.EOF = False Then
            m_typDetail_Rec(intIndex1).F = adoRecordset1.RecordCount
        Else
            m_typDetail_Rec(intIndex1).F = 0
        End If
        adoRecordset1.Close
    
        If m_typDetail_Rec(intIndex1).Num = 0 Then
            '�������׃f�[�^
            strSQL = "SELECT * FROM DT031" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode & _
                     " AND (Bnum IS NULL OR Bnum = " & m_typDetail_Rec(intIndex1).Num & ")"
        Else
            '�������׃f�[�^
            strSQL = "SELECT * FROM DT031" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode & _
                     " AND Bnum = " & m_typDetail_Rec(intIndex1).Num
        End If
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoRecordset1.EOF = False Then
            m_typDetail_Rec(intIndex1).F = m_typDetail_Rec(intIndex1).F + CInt(adoRecordset1.RecordCount)
        End If
        adoRecordset1.Close
    Next intIndex1
    
    Exit Sub
    
Get_F_Err:

    Call MsgBox("F�����̎擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Get_F_Err")

End Sub

'�ځ@�I�@�@�FActiveReport�̈��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F0:�v���r���[ 1:���
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�O�W
'�X�V�����@�F
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf100
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "�����ꗗ�\"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .Caption = "�����ꗗ�\"
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

'�ځ@�I�@�@�F����p���[�N�쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�Q�Q
'�X�V�����@�F
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim wkRecordset As New ADODB.Recordset
    Dim lngIndex1 As Long
    Dim strBuff1 As String


    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    If UBound(m_typDetail_Rec) <= 0 Then
        Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF100"
    g_clsAdoAccess.Connection.Execute strSQL

    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF100"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    With frmCount
        .fpProgressBar1.Value = 0
        .fpProgressBar1.Max = UBound(m_typDetail_Rec)
        .Show
        Me.Enabled = False
    End With
    
    For lngIndex1 = 1 To UBound(m_typDetail_Rec)
        If m_typDetail_Rec(lngIndex1).Zandaka <> 0 Then
            wkRecordset.AddNew
            wkRecordset.Fields("Odate") = lblOdate.Caption
            wkRecordset.Fields("Bcode") = m_typDetail_Rec(lngIndex1).Bcode
            wkRecordset.Fields("Bname") = m_typDetail_Rec(lngIndex1).Bname
            wkRecordset.Fields("Total") = m_typDetail_Rec(lngIndex1).Total
            wkRecordset.Fields("Tax") = m_typDetail_Rec(lngIndex1).Tax
            wkRecordset.Fields("Price") = m_typDetail_Rec(lngIndex1).Total2
            wkRecordset.Fields("Keep") = m_typDetail_Rec(lngIndex1).Keep
            wkRecordset.Fields("GTotal") = m_typDetail_Rec(lngIndex1).Gtotal
            '201107
            wkRecordset.Fields("Brate2") = m_typDetail_Rec(lngIndex1).Brate2
            wkRecordset.Update
        End If
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Next lngIndex1
    
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

'�ځ@�I�@�@�F�R���{�{�b�N�X�̍쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�O
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
