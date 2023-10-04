VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmYpmf060 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   8835
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12570
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf060.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   12570
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraInput 
      Height          =   1635
      Left            =   60
      TabIndex        =   127
      Top             =   1380
      Width           =   12435
      Begin VB.CommandButton cmdFuriwake 
         Caption         =   "↓明細振分↓"
         Height          =   375
         Left            =   3540
         TabIndex        =   14
         Top             =   1140
         Width           =   1515
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   128
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "入金種別"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin imNumber6Ctl.imNumber imnInpRdate_Year 
         Height          =   375
         Left            =   1620
         TabIndex        =   10
         Top             =   660
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         Calculator      =   "frmYpmf060.frx":0CFA
         Caption         =   "frmYpmf060.frx":0D1A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf060.frx":0D88
         Keys            =   "frmYpmf060.frx":0DA6
         Spin            =   "frmYpmf060.frx":0DE0
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
      Begin imNumber6Ctl.imNumber imnInp_Rdate_Month 
         Height          =   375
         Left            =   2940
         TabIndex        =   11
         Top             =   660
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   661
         Calculator      =   "frmYpmf060.frx":0E08
         Caption         =   "frmYpmf060.frx":0E28
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf060.frx":0E96
         Keys            =   "frmYpmf060.frx":0EB4
         Spin            =   "frmYpmf060.frx":0EEE
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
      Begin imNumber6Ctl.imNumber imnInp_Rdate_Day 
         Height          =   375
         Left            =   3900
         TabIndex        =   12
         Top             =   660
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   661
         Calculator      =   "frmYpmf060.frx":0F16
         Caption         =   "frmYpmf060.frx":0F36
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf060.frx":0FA4
         Keys            =   "frmYpmf060.frx":0FC2
         Spin            =   "frmYpmf060.frx":0FFC
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
      Begin CSComboLib.CSComboBox cboInpR 
         Height          =   405
         Left            =   1620
         TabIndex        =   9
         Top             =   180
         Width           =   1635
         _Version        =   262145
         _ExtentX        =   2884
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         Contents        =   "frmYpmf060.frx":1024
         Text            =   "銀行振込"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   132
         Top             =   660
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "入金日"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         Index           =   3
         Left            =   120
         TabIndex        =   133
         Top             =   1140
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "入金額"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin imNumber6Ctl.imNumber imnInp_Nyukin 
         Height          =   375
         Left            =   1620
         TabIndex        =   13
         Top             =   1140
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   661
         Calculator      =   "frmYpmf060.frx":103D
         Caption         =   "frmYpmf060.frx":105D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf060.frx":10CB
         Keys            =   "frmYpmf060.frx":10E9
         Spin            =   "frmYpmf060.frx":1133
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
         ValueVT         =   1245189
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   2580
         TabIndex        =   131
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   3540
         TabIndex        =   130
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   4500
         TabIndex        =   129
         Top             =   720
         Width           =   315
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   60
      TabIndex        =   113
      Top             =   660
      Width           =   12435
      Begin VB.CommandButton cmdReset 
         Caption         =   "入金取消画面表示"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9540
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   2355
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "検索開始"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7200
         TabIndex        =   6
         Top             =   180
         Width           =   1635
      End
      Begin CSComboLib.CSComboBox cboBcode 
         Height          =   405
         Index           =   0
         Left            =   1620
         TabIndex        =   4
         Top             =   180
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf060.frx":115B
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   114
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "買主ｺｰﾄﾞ抽出"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf060.frx":1174
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin VB.Label Label2 
         Alignment       =   2  '中央揃え
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7140
         TabIndex        =   117
         Top             =   180
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblScode_Name 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   116
         Top             =   180
         Visible         =   0   'False
         Width           =   4275
      End
      Begin VB.Label lblScode_Name 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   115
         Top             =   180
         Width           =   4395
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   9840
      TabIndex        =   112
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.OptionButton optDisplayFlg 
         Caption         =   "全データ表示"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optDisplayFlg 
         Caption         =   "未入金のみ表示"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   88
      Top             =   7380
      Width           =   12435
      Begin imNumber6Ctl.imNumber imnGtotal_Total 
         Height          =   375
         Left            =   4860
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   661
         Calculator      =   "frmYpmf060.frx":118D
         Caption         =   "frmYpmf060.frx":11AD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf060.frx":121B
         Keys            =   "frmYpmf060.frx":1239
         Spin            =   "frmYpmf060.frx":1283
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
         ValueVT         =   2011496453
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnZandaka_Total 
         Height          =   375
         Left            =   6180
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   661
         Calculator      =   "frmYpmf060.frx":12AB
         Caption         =   "frmYpmf060.frx":12CB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf060.frx":1339
         Keys            =   "frmYpmf060.frx":1357
         Spin            =   "frmYpmf060.frx":13A1
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
         Left            =   10380
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   180
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1014
         _ExtentY        =   661
         Calculator      =   "frmYpmf060.frx":13C9
         Caption         =   "frmYpmf060.frx":13E9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf060.frx":1457
         Keys            =   "frmYpmf060.frx":1475
         Spin            =   "frmYpmf060.frx":14BF
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
      Begin imNumber6Ctl.imNumber imnBrate2_Total 
         Height          =   375
         Left            =   10940
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
         _Version        =   65536
         _ExtentX        =   1640
         _ExtentY        =   661
         Calculator      =   "frmYpmf060.frx":14E7
         Caption         =   "frmYpmf060.frx":1507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf060.frx":1575
         Keys            =   "frmYpmf060.frx":1593
         Spin            =   "frmYpmf060.frx":15DD
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
   End
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   63
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton cmdLogin 
         Caption         =   "開催年月日と担当者の変更"
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9960
         TabIndex        =   64
         Top             =   180
         Visible         =   0   'False
         Width           =   675
         _Version        =   262145
         _ExtentX        =   1191
         _ExtentY        =   635
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "2;20"
         Contents        =   "frmYpmf060.frx":1605
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   67
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "開催年月日"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   68
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "担当者"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "ＮＮＮＮＮＮＮＮＮＮ"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   66
         Top             =   180
         Width           =   1905
      End
      Begin VB.Label lblOdate 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "9999/12/31"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   65
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   62
      Top             =   7980
      Width           =   12435
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   57
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "画面クリア(F8)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         rText.left      =   4
         rText.top       =   8
         rText.right     =   109
         rText.bottom    =   27
         Picture         =   "frmYpmf060.frx":161E
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10020
         TabIndex        =   56
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "終了(F9)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         Picture         =   "frmYpmf060.frx":163A
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   1860
         TabIndex        =   58
         Top             =   180
         Width           =   2235
         _Version        =   262145
         _ExtentX        =   3942
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "入金一覧表F10)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         rText.left      =   26
         rText.top       =   8
         rText.right     =   143
         rText.bottom    =   27
         Picture         =   "frmYpmf060.frx":1794
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdInput 
         Height          =   495
         Left            =   8220
         TabIndex        =   55
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "入金(F12)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         Picture         =   "frmYpmf060.frx":18A6
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   4275
      Left            =   60
      TabIndex        =   61
      Top             =   3060
      Width           =   12435
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   85
         Top             =   3480
         Width           =   11835
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   5
            Left            =   10820
            TabIndex        =   150
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":1CF8
            Caption         =   "frmYpmf060.frx":1D18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":1D86
            Keys            =   "frmYpmf060.frx":1DA4
            Spin            =   "frmYpmf060.frx":1DEE
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.CommandButton cmdPayment 
            Caption         =   "入　金"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   13440
            TabIndex        =   53
            Top             =   180
            Visible         =   0   'False
            Width           =   795
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   5
            Left            =   4740
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":1E16
            Caption         =   "frmYpmf060.frx":1E36
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":1EA4
            Keys            =   "frmYpmf060.frx":1EC2
            Spin            =   "frmYpmf060.frx":1F0C
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
            ValueVT         =   2011496453
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Month 
            Height          =   375
            Index           =   5
            Left            =   11940
            TabIndex        =   51
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":1F34
            Caption         =   "frmYpmf060.frx":1F54
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":1FC2
            Keys            =   "frmYpmf060.frx":1FE0
            Spin            =   "frmYpmf060.frx":201A
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
            ValueVT         =   2088828933
            Value           =   12
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Day 
            Height          =   375
            Index           =   5
            Left            =   12660
            TabIndex        =   52
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":2042
            Caption         =   "frmYpmf060.frx":2062
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":20D0
            Keys            =   "frmYpmf060.frx":20EE
            Spin            =   "frmYpmf060.frx":2128
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
            ValueVT         =   2088828933
            Value           =   31
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   5
            Left            =   6060
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":2150
            Caption         =   "frmYpmf060.frx":2170
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":21DE
            Keys            =   "frmYpmf060.frx":21FC
            Spin            =   "frmYpmf060.frx":2246
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
            ValueVT         =   2011496453
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnNyukin 
            Height          =   375
            Index           =   5
            Left            =   8890
            TabIndex        =   47
            Top             =   180
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":226E
            Caption         =   "frmYpmf060.frx":228E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":22FC
            Keys            =   "frmYpmf060.frx":231A
            Spin            =   "frmYpmf060.frx":2364
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
            ReadOnly        =   1
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
            Left            =   10260
            TabIndex        =   48
            Top             =   180
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":238C
            Caption         =   "frmYpmf060.frx":23AC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":241A
            Keys            =   "frmYpmf060.frx":2438
            Spin            =   "frmYpmf060.frx":2482
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeepZan 
            Height          =   375
            Index           =   5
            Left            =   7380
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   180
            Width           =   575
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":24AA
            Caption         =   "frmYpmf060.frx":24CA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":2538
            Keys            =   "frmYpmf060.frx":2556
            Spin            =   "frmYpmf060.frx":25A0
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
         Begin imNumber6Ctl.imNumber imnBrate2Zan 
            Height          =   375
            Index           =   5
            Left            =   7940
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":25C8
            Caption         =   "frmYpmf060.frx":25E8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":2656
            Keys            =   "frmYpmf060.frx":2674
            Spin            =   "frmYpmf060.frx":26BE
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
         Begin imNumber6Ctl.imNumber imnRdate_Year 
            Height          =   375
            Index           =   5
            Left            =   10920
            TabIndex        =   50
            Top             =   180
            Visible         =   0   'False
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":26E6
            Caption         =   "frmYpmf060.frx":2706
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":2774
            Keys            =   "frmYpmf060.frx":2792
            Spin            =   "frmYpmf060.frx":27CC
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
            ValueVT         =   5
            Value           =   2099
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSComboLib.CSComboBox cboR 
            Height          =   315
            Index           =   5
            Left            =   10680
            TabIndex        =   49
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.76
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColDelim        =   ";"
            Contents        =   "frmYpmf060.frx":27F4
            Text            =   "銀行振込"
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "日"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   13140
            TabIndex        =   109
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   12360
            TabIndex        =   108
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   11640
            TabIndex        =   107
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblOdate_Detail 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999/12/31"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   3420
            TabIndex        =   94
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
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
            TabIndex        =   87
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblBname 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   720
            TabIndex        =   86
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   82
         Top             =   2820
         Width           =   11835
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   4
            Left            =   10820
            TabIndex        =   148
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":280D
            Caption         =   "frmYpmf060.frx":282D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":289B
            Keys            =   "frmYpmf060.frx":28B9
            Spin            =   "frmYpmf060.frx":2903
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.CommandButton cmdPayment 
            Caption         =   "入　金"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   13440
            TabIndex        =   45
            Top             =   180
            Visible         =   0   'False
            Width           =   795
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   4
            Left            =   4740
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":292B
            Caption         =   "frmYpmf060.frx":294B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":29B9
            Keys            =   "frmYpmf060.frx":29D7
            Spin            =   "frmYpmf060.frx":2A21
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
            ValueVT         =   2011496453
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Month 
            Height          =   375
            Index           =   4
            Left            =   11940
            TabIndex        =   43
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":2A49
            Caption         =   "frmYpmf060.frx":2A69
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":2AD7
            Keys            =   "frmYpmf060.frx":2AF5
            Spin            =   "frmYpmf060.frx":2B2F
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
            ValueVT         =   2088828933
            Value           =   12
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Day 
            Height          =   375
            Index           =   4
            Left            =   12660
            TabIndex        =   44
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":2B57
            Caption         =   "frmYpmf060.frx":2B77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":2BE5
            Keys            =   "frmYpmf060.frx":2C03
            Spin            =   "frmYpmf060.frx":2C3D
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
            ValueVT         =   2088828933
            Value           =   31
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   4
            Left            =   6060
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":2C65
            Caption         =   "frmYpmf060.frx":2C85
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":2CF3
            Keys            =   "frmYpmf060.frx":2D11
            Spin            =   "frmYpmf060.frx":2D5B
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
            ValueVT         =   2011496453
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnNyukin 
            Height          =   375
            Index           =   4
            Left            =   8890
            TabIndex        =   39
            Top             =   180
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":2D83
            Caption         =   "frmYpmf060.frx":2DA3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":2E11
            Keys            =   "frmYpmf060.frx":2E2F
            Spin            =   "frmYpmf060.frx":2E79
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
            ReadOnly        =   1
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
            Left            =   10260
            TabIndex        =   40
            Top             =   180
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":2EA1
            Caption         =   "frmYpmf060.frx":2EC1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":2F2F
            Keys            =   "frmYpmf060.frx":2F4D
            Spin            =   "frmYpmf060.frx":2F97
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeepZan 
            Height          =   375
            Index           =   4
            Left            =   7380
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   180
            Width           =   575
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":2FBF
            Caption         =   "frmYpmf060.frx":2FDF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":304D
            Keys            =   "frmYpmf060.frx":306B
            Spin            =   "frmYpmf060.frx":30B5
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
         Begin imNumber6Ctl.imNumber imnBrate2Zan 
            Height          =   375
            Index           =   4
            Left            =   7940
            TabIndex        =   147
            TabStop         =   0   'False
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":30DD
            Caption         =   "frmYpmf060.frx":30FD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":316B
            Keys            =   "frmYpmf060.frx":3189
            Spin            =   "frmYpmf060.frx":31D3
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
         Begin imNumber6Ctl.imNumber imnRdate_Year 
            Height          =   375
            Index           =   4
            Left            =   10920
            TabIndex        =   42
            Top             =   180
            Visible         =   0   'False
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":31FB
            Caption         =   "frmYpmf060.frx":321B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3289
            Keys            =   "frmYpmf060.frx":32A7
            Spin            =   "frmYpmf060.frx":32E1
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
            ValueVT         =   5
            Value           =   2099
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSComboLib.CSComboBox cboR 
            Height          =   315
            Index           =   4
            Left            =   10680
            TabIndex        =   41
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.76
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColDelim        =   ";"
            Contents        =   "frmYpmf060.frx":3309
            Text            =   "銀行振込"
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "日"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   13140
            TabIndex        =   106
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   12360
            TabIndex        =   105
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   11640
            TabIndex        =   104
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblOdate_Detail 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999/12/31"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   3420
            TabIndex        =   93
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
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
            TabIndex        =   84
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblBname 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   720
            TabIndex        =   83
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   79
         Top             =   2160
         Width           =   11835
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   3
            Left            =   10820
            TabIndex        =   146
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":3322
            Caption         =   "frmYpmf060.frx":3342
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":33B0
            Keys            =   "frmYpmf060.frx":33CE
            Spin            =   "frmYpmf060.frx":3418
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.CommandButton cmdPayment 
            Caption         =   "入　金"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   13440
            TabIndex        =   37
            Top             =   180
            Visible         =   0   'False
            Width           =   795
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   3
            Left            =   4740
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":3440
            Caption         =   "frmYpmf060.frx":3460
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":34CE
            Keys            =   "frmYpmf060.frx":34EC
            Spin            =   "frmYpmf060.frx":3536
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
            ValueVT         =   2011496453
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Month 
            Height          =   375
            Index           =   3
            Left            =   11940
            TabIndex        =   35
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":355E
            Caption         =   "frmYpmf060.frx":357E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":35EC
            Keys            =   "frmYpmf060.frx":360A
            Spin            =   "frmYpmf060.frx":3644
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
            ValueVT         =   2088828933
            Value           =   12
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Day 
            Height          =   375
            Index           =   3
            Left            =   12660
            TabIndex        =   36
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":366C
            Caption         =   "frmYpmf060.frx":368C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":36FA
            Keys            =   "frmYpmf060.frx":3718
            Spin            =   "frmYpmf060.frx":3752
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
            ValueVT         =   2088828933
            Value           =   31
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   3
            Left            =   6060
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":377A
            Caption         =   "frmYpmf060.frx":379A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3808
            Keys            =   "frmYpmf060.frx":3826
            Spin            =   "frmYpmf060.frx":3870
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
            ValueVT         =   2011496453
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnNyukin 
            Height          =   375
            Index           =   3
            Left            =   8890
            TabIndex        =   31
            Top             =   180
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":3898
            Caption         =   "frmYpmf060.frx":38B8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3926
            Keys            =   "frmYpmf060.frx":3944
            Spin            =   "frmYpmf060.frx":398E
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
            ReadOnly        =   1
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
            Left            =   10260
            TabIndex        =   32
            Top             =   180
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":39B6
            Caption         =   "frmYpmf060.frx":39D6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3A44
            Keys            =   "frmYpmf060.frx":3A62
            Spin            =   "frmYpmf060.frx":3AAC
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeepZan 
            Height          =   375
            Index           =   3
            Left            =   7380
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   180
            Width           =   575
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":3AD4
            Caption         =   "frmYpmf060.frx":3AF4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3B62
            Keys            =   "frmYpmf060.frx":3B80
            Spin            =   "frmYpmf060.frx":3BCA
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
         Begin imNumber6Ctl.imNumber imnBrate2Zan 
            Height          =   375
            Index           =   3
            Left            =   7940
            TabIndex        =   145
            TabStop         =   0   'False
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":3BF2
            Caption         =   "frmYpmf060.frx":3C12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3C80
            Keys            =   "frmYpmf060.frx":3C9E
            Spin            =   "frmYpmf060.frx":3CE8
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
         Begin imNumber6Ctl.imNumber imnRdate_Year 
            Height          =   375
            Index           =   3
            Left            =   10920
            TabIndex        =   34
            Top             =   180
            Visible         =   0   'False
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":3D10
            Caption         =   "frmYpmf060.frx":3D30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3D9E
            Keys            =   "frmYpmf060.frx":3DBC
            Spin            =   "frmYpmf060.frx":3DF6
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
            ValueVT         =   5
            Value           =   2099
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSComboLib.CSComboBox cboR 
            Height          =   315
            Index           =   3
            Left            =   10680
            TabIndex        =   33
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.76
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColDelim        =   ";"
            Contents        =   "frmYpmf060.frx":3E1E
            Text            =   "銀行振込"
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "日"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   13140
            TabIndex        =   103
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   12360
            TabIndex        =   102
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Left            =   11640
            TabIndex        =   101
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblOdate_Detail 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999/12/31"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   3420
            TabIndex        =   92
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
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
            TabIndex        =   81
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblBname 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   720
            TabIndex        =   80
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   76
         Top             =   1500
         Width           =   11835
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   2
            Left            =   10820
            TabIndex        =   144
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":3E37
            Caption         =   "frmYpmf060.frx":3E57
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3EC5
            Keys            =   "frmYpmf060.frx":3EE3
            Spin            =   "frmYpmf060.frx":3F2D
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.CommandButton cmdPayment 
            Caption         =   "入　金"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   13440
            TabIndex        =   29
            Top             =   180
            Visible         =   0   'False
            Width           =   795
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   2
            Left            =   4740
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":3F55
            Caption         =   "frmYpmf060.frx":3F75
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":3FE3
            Keys            =   "frmYpmf060.frx":4001
            Spin            =   "frmYpmf060.frx":404B
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
            ValueVT         =   2011496453
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Month 
            Height          =   375
            Index           =   2
            Left            =   11940
            TabIndex        =   27
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4073
            Caption         =   "frmYpmf060.frx":4093
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4101
            Keys            =   "frmYpmf060.frx":411F
            Spin            =   "frmYpmf060.frx":4159
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
            ValueVT         =   2088828933
            Value           =   12
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Day 
            Height          =   375
            Index           =   2
            Left            =   12660
            TabIndex        =   28
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4181
            Caption         =   "frmYpmf060.frx":41A1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":420F
            Keys            =   "frmYpmf060.frx":422D
            Spin            =   "frmYpmf060.frx":4267
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
            ValueVT         =   2088828933
            Value           =   31
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnZandaka 
            Height          =   375
            Index           =   2
            Left            =   6060
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":428F
            Caption         =   "frmYpmf060.frx":42AF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":431D
            Keys            =   "frmYpmf060.frx":433B
            Spin            =   "frmYpmf060.frx":4385
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
            ValueVT         =   2011496453
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnNyukin 
            Height          =   375
            Index           =   2
            Left            =   8890
            TabIndex        =   23
            Top             =   180
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":43AD
            Caption         =   "frmYpmf060.frx":43CD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":443B
            Keys            =   "frmYpmf060.frx":4459
            Spin            =   "frmYpmf060.frx":44A3
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
            ReadOnly        =   1
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
            Left            =   10260
            TabIndex        =   24
            Top             =   180
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":44CB
            Caption         =   "frmYpmf060.frx":44EB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4559
            Keys            =   "frmYpmf060.frx":4577
            Spin            =   "frmYpmf060.frx":45C1
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeepZan 
            Height          =   375
            Index           =   2
            Left            =   7380
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   180
            Width           =   575
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":45E9
            Caption         =   "frmYpmf060.frx":4609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4677
            Keys            =   "frmYpmf060.frx":4695
            Spin            =   "frmYpmf060.frx":46DF
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
         Begin imNumber6Ctl.imNumber imnBrate2Zan 
            Height          =   375
            Index           =   2
            Left            =   7940
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4707
            Caption         =   "frmYpmf060.frx":4727
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4795
            Keys            =   "frmYpmf060.frx":47B3
            Spin            =   "frmYpmf060.frx":47FD
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
         Begin imNumber6Ctl.imNumber imnRdate_Year 
            Height          =   375
            Index           =   2
            Left            =   10920
            TabIndex        =   26
            Top             =   180
            Visible         =   0   'False
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4825
            Caption         =   "frmYpmf060.frx":4845
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":48B3
            Keys            =   "frmYpmf060.frx":48D1
            Spin            =   "frmYpmf060.frx":490B
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
            ValueVT         =   5
            Value           =   2099
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSComboLib.CSComboBox cboR 
            Height          =   315
            Index           =   2
            Left            =   10680
            TabIndex        =   25
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.76
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColDelim        =   ";"
            Contents        =   "frmYpmf060.frx":4933
            Text            =   "銀行振込"
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "日"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Left            =   13140
            TabIndex        =   100
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   12360
            TabIndex        =   99
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Left            =   11640
            TabIndex        =   98
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblOdate_Detail 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999/12/31"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   3420
            TabIndex        =   91
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
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
            TabIndex        =   78
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblBname 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   77
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3135
         Left            =   12000
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   960
         Width           =   315
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   72
         Top             =   180
         Width           =   11835
         Begin VB.Label Label1 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "(維費)     (手料)"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   21
            Left            =   7200
            TabIndex        =   139
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "(維費)   (手料)"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   37
            Left            =   10320
            TabIndex        =   126
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "今回入金額"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   36
            Left            =   9000
            TabIndex        =   125
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "売上合計"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   4560
            TabIndex        =   119
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "開催日"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   34
            Left            =   3420
            TabIndex        =   111
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "入金種別"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   11820
            TabIndex        =   110
            Top             =   240
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label Label1 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "入金日"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Left            =   13080
            TabIndex        =   89
            Top             =   240
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "残高金額"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Left            =   5880
            TabIndex        =   74
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label1 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "買　主"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            TabIndex        =   73
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   69
         Top             =   840
         Width           =   11835
         Begin imNumber6Ctl.imNumber imnBrate2 
            Height          =   375
            Index           =   1
            Left            =   10820
            TabIndex        =   142
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":494C
            Caption         =   "frmYpmf060.frx":496C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":49DA
            Keys            =   "frmYpmf060.frx":49F8
            Spin            =   "frmYpmf060.frx":4A42
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.CommandButton cmdPayment 
            Caption         =   "入　金"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   13440
            TabIndex        =   21
            Top             =   180
            Visible         =   0   'False
            Width           =   795
         End
         Begin imNumber6Ctl.imNumber imnGtotal 
            Height          =   375
            Index           =   1
            Left            =   4740
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4A6A
            Caption         =   "frmYpmf060.frx":4A8A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4AF8
            Keys            =   "frmYpmf060.frx":4B16
            Spin            =   "frmYpmf060.frx":4B60
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
            ValueVT         =   2088828933
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Month 
            Height          =   375
            Index           =   1
            Left            =   11940
            TabIndex        =   19
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4B88
            Caption         =   "frmYpmf060.frx":4BA8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4C16
            Keys            =   "frmYpmf060.frx":4C34
            Spin            =   "frmYpmf060.frx":4C6E
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
            ValueVT         =   5
            Value           =   12
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRdate_Day 
            Height          =   375
            Index           =   1
            Left            =   12660
            TabIndex        =   20
            Top             =   180
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4C96
            Caption         =   "frmYpmf060.frx":4CB6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4D24
            Keys            =   "frmYpmf060.frx":4D42
            Spin            =   "frmYpmf060.frx":4D7C
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
            ValueVT         =   2088828933
            Value           =   31
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnNyukin 
            Height          =   375
            Index           =   1
            Left            =   8890
            TabIndex        =   15
            Top             =   180
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4DA4
            Caption         =   "frmYpmf060.frx":4DC4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4E32
            Keys            =   "frmYpmf060.frx":4E50
            Spin            =   "frmYpmf060.frx":4E9A
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
            ReadOnly        =   1
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
            Left            =   6060
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4EC2
            Caption         =   "frmYpmf060.frx":4EE2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":4F50
            Keys            =   "frmYpmf060.frx":4F6E
            Spin            =   "frmYpmf060.frx":4FB8
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
            Left            =   10260
            TabIndex        =   16
            Top             =   180
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":4FE0
            Caption         =   "frmYpmf060.frx":5000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":506E
            Keys            =   "frmYpmf060.frx":508C
            Spin            =   "frmYpmf060.frx":50D6
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
            ReadOnly        =   1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnKeepZan 
            Height          =   375
            Index           =   1
            Left            =   7380
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   180
            Width           =   575
            _Version        =   65536
            _ExtentX        =   1014
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":50FE
            Caption         =   "frmYpmf060.frx":511E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":518C
            Keys            =   "frmYpmf060.frx":51AA
            Spin            =   "frmYpmf060.frx":51F4
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
         Begin imNumber6Ctl.imNumber imnBrate2Zan 
            Height          =   375
            Index           =   1
            Left            =   7940
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   180
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":521C
            Caption         =   "frmYpmf060.frx":523C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":52AA
            Keys            =   "frmYpmf060.frx":52C8
            Spin            =   "frmYpmf060.frx":5312
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
         Begin CSComboLib.CSComboBox cboR 
            Height          =   315
            Index           =   1
            Left            =   9240
            TabIndex        =   17
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.76
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColDelim        =   ";"
            Contents        =   "frmYpmf060.frx":533A
            Text            =   "銀行振込"
         End
         Begin imNumber6Ctl.imNumber imnRdate_Year 
            Height          =   375
            Index           =   1
            Left            =   9240
            TabIndex        =   18
            Top             =   180
            Visible         =   0   'False
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   661
            Calculator      =   "frmYpmf060.frx":5353
            Caption         =   "frmYpmf060.frx":5373
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf060.frx":53E1
            Keys            =   "frmYpmf060.frx":53FF
            Spin            =   "frmYpmf060.frx":5439
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
            ValueVT         =   5
            Value           =   2099
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "日"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Left            =   13140
            TabIndex        =   97
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Left            =   12360
            TabIndex        =   96
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   11640
            TabIndex        =   95
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblOdate_Detail 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999/12/31"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   3420
            TabIndex        =   90
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label lblBname 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   71
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblBcode 
            Alignment       =   1  '右揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
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
            TabIndex        =   70
            Top             =   180
            Width           =   615
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
      Caption         =   "frmYpmf060.frx":5461
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf060.frx":54CF
      Key             =   "frmYpmf060.frx":54ED
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
      TabIndex        =   59
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf060.frx":5531
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf060.frx":559F
      Key             =   "frmYpmf060.frx":55BD
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
      TabIndex        =   60
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf060.frx":5601
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf060.frx":566F
      Key             =   "frmYpmf060.frx":568D
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
Attribute VB_Name = "frmYpmf060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DETAIL_MAX = 5               '明細表示行数

Private Const BACK_COLOR_ON = &HFF0000
Private Const BACK_COLOR_OFF = &H8000000F

Private Const CMD_PAYMENT_OFF = "入　金"
Private Const CMD_PAYMENT_ON = "入金済"

Private Const PAYMENT_DIV1_NAME = "現金"
Private Const PAYMENT_DIV2_NAME = "小切手"
Private Const PAYMENT_DIV3_NAME = "銀行振込"

Private Type Detail_Record
    Bcode As String
    Bname As String
    Odate As String
    Gtotal As Currency
    Zandaka As Currency
    KeepZan As Currency
    Brate2Zan As Currency
    Nyukin As Currency
    Keep As Currency
    Brate2 As Currency
    Rdiv As Integer
    R As Integer
    Rdate As String
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
        '得意先マスタ
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

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBcode_Validate_Err")

End Sub

Private Sub cboInpR_GotFocus()

    cboInpR.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub cboInpR_LostFocus()
    
    cboInpR.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub cboR_GotFocus(Index As Integer)

    cboR(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub cboR_LostFocus(Index As Integer)

    cboR(Index).BackColor = FOCUS_NO_COLOR

End Sub

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)

End Sub

Private Sub cmdExecute_Click()

    frmPrintDialog.Show vbModal

End Sub

'目　的　　：
'条　件　　：終了クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdFuriwake_Click()

    Dim intIndex1 As Integer
    Dim curNyukin As Currency

    On Error GoTo cmdFuriwake_Click_Err
            
    '入力チェック
    If Trim(cboBcode(0).Text) = "" Then
        MsgBox "買主コードが入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If cboInpR.Text = "" Then
        MsgBox "入金種別が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If Trim(imnInpRdate_Year.Text) = "" Then
        MsgBox "入金日が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If Trim(imnInp_Rdate_Month.Text) = "" Then
        MsgBox "入金日が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If Trim(imnInp_Rdate_Day.Text) = "" Then
        MsgBox "入金日が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If Global_IsDate(imnInpRdate_Year.Text, imnInp_Rdate_Month.Text, imnInp_Rdate_Day.Text) = False Then
        MsgBox "正しい入金日が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If imnInp_Nyukin.Value = 0 Then
        MsgBox "今回入金額が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
            
    curNyukin = imnInp_Nyukin.Value
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
        If m_typDetail_Rec(intIndex1).Zandaka <= curNyukin Then
            m_typDetail_Rec(intIndex1).Nyukin = m_typDetail_Rec(intIndex1).Zandaka
            m_typDetail_Rec(intIndex1).Keep = m_typDetail_Rec(intIndex1).KeepZan
            '201107
            m_typDetail_Rec(intIndex1).Brate2 = m_typDetail_Rec(intIndex1).Brate2Zan
        
            curNyukin = curNyukin - m_typDetail_Rec(intIndex1).Zandaka
        Else
            m_typDetail_Rec(intIndex1).Nyukin = curNyukin
            m_typDetail_Rec(intIndex1).Keep = 0
                    
            curNyukin = 0
        End If
    Next intIndex1
    
    Call Detail_Dislplay(1)
    Call Detail_ScrollBar
    Call Calc_Total
    
    Exit Sub

cmdFuriwake_Click_Err:

    Call MsgBox("明細振分クリック時エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdFuriwake_Click_Err")

End Sub

Private Sub cmdInput_Click()
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strRdate As String
    Dim intR As Integer
    Dim intIndex1 As Integer
    Dim blnChk As Boolean
    
    On Error GoTo cmdInput_Click_Err
    
    If MsgBox("入金処理しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    '入力チェック
    If Trim(cboBcode(0).Text) = "" Then
        MsgBox "買主コードが入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If cboInpR.Text = "" Then
        MsgBox "入金種別が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If Trim(imnInpRdate_Year.Text) = "" Then
        MsgBox "入金日が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If Trim(imnInp_Rdate_Month.Text) = "" Then
        MsgBox "入金日が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If Trim(imnInp_Rdate_Day.Text) = "" Then
        MsgBox "入金日が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If Global_IsDate(imnInpRdate_Year.Text, imnInp_Rdate_Month.Text, imnInp_Rdate_Day.Text) = False Then
        MsgBox "正しい入金日が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    If imnInp_Nyukin.Value = 0 Then
        MsgBox "今回入金額が入力されていません。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If
    
    '入金振分を行っているかチェック
    blnChk = False
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
        If m_typDetail_Rec(intIndex1).Nyukin <> 0 Then
            blnChk = True
            Exit For
        End If
    Next intIndex1
    If blnChk = False Then
        MsgBox "明細振分をクリックしてください。", vbOKOnly + vbCritical, "入力チェック"
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    
    '入金種別
    Select Case Trim(cboInpR.Text)
        Case PAYMENT_DIV1_NAME:
            intR = PAYMENT_DIV_CASH
        Case PAYMENT_DIV2_NAME:
            intR = PAYMENT_DIV_CHECK
        Case PAYMENT_DIV3_NAME:
            intR = PAYMENT_DIV_TRANSFER
        Case Else
            intR = PAYMENT_DIV_CASH
    End Select
    
    '入金日
    strRdate = Format(imnInpRdate_Year.Text, "0000") & "/" & Format(imnInp_Rdate_Month.Text, "00") & "/" & Format(imnInp_Rdate_Day.Text, "00")
    
    g_clsAdoSQL.Connection.BeginTrans
    
    '明細分登録する
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
        If m_typDetail_Rec(intIndex1).Nyukin <> 0 Or _
           m_typDetail_Rec(intIndex1).Keep <> 0 Then
            
            '入金データ
            strSQL = "SELECT * FROM DT060" & _
                     " WHERE Odate = '" & m_typDetail_Rec(intIndex1).Odate & "'" & _
                     " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode & _
                     " AND Rdate = '" & strRdate & "'" & _
                     " ORDER BY Odate,Bcode,Rdate"
            adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If adoRecordset1.EOF = True Then
                adoRecordset1.AddNew
                adoRecordset1.Fields("Odate") = m_typDetail_Rec(intIndex1).Odate
                adoRecordset1.Fields("Bcode") = m_typDetail_Rec(intIndex1).Bcode
                adoRecordset1.Fields("Rdate") = strRdate
                adoRecordset1.Fields("Rdiv") = PAYMENT_ON
                adoRecordset1.Fields("R") = intR
                adoRecordset1.Fields("Ptotal") = m_typDetail_Rec(intIndex1).Nyukin
                adoRecordset1.Fields("Ptotal2") = m_typDetail_Rec(intIndex1).Keep
                '201107
                adoRecordset1.Fields("Ptotal3") = m_typDetail_Rec(intIndex1).Brate2
                adoRecordset1.Update
            Else
                '入金日が同じデータがあった場合、加算する
                adoRecordset1.Fields("Rdiv") = PAYMENT_ON
                adoRecordset1.Fields("R") = intR
                adoRecordset1.Fields("Ptotal") = CCur(adoRecordset1.Fields("Ptotal")) + m_typDetail_Rec(intIndex1).Nyukin
                adoRecordset1.Fields("Ptotal2") = CCur(adoRecordset1.Fields("Ptotal2")) + m_typDetail_Rec(intIndex1).Keep
                '201107
                adoRecordset1.Fields("Ptotal3") = CCur(adoRecordset1.Fields("Ptotal3")) + m_typDetail_Rec(intIndex1).Brate2
                adoRecordset1.Update
            End If
            adoRecordset1.Close
    
            '残高がゼロになった場合のみ買主精算データを更新する
            If (m_typDetail_Rec(intIndex1).Zandaka - m_typDetail_Rec(intIndex1).Nyukin) <= 0 Then
                
                m_typDetail_Rec(intIndex1).Zandaka = 0
                m_typDetail_Rec(intIndex1).KeepZan = 0
                m_typDetail_Rec(intIndex1).Nyukin = 0
                m_typDetail_Rec(intIndex1).Keep = 0
                m_typDetail_Rec(intIndex1).Rdiv = PAYMENT_ON
                '201107
                m_typDetail_Rec(intIndex1).Brate2 = 0
                m_typDetail_Rec(intIndex1).Brate2Zan = 0
                
                '買主精算データ
                strSQL = "SELECT * FROM DT041" & _
                         " WHERE Odate = '" & m_typDetail_Rec(intIndex1).Odate & "'" & _
                         " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode & _
                         " ORDER BY Bcode,Odate,Num"
                adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
                Do While Not adoRecordset1.EOF
                    adoRecordset1.Fields("Rdiv") = PAYMENT_ON
                    adoRecordset1.Fields("R") = intR
                    adoRecordset1.Fields("Rdate") = strRdate
                    adoRecordset1.Update
                
                    adoRecordset1.MoveNext
                Loop
                adoRecordset1.Close
            
                '買主精算データ(累積)
                strSQL = "SELECT * FROM RT041" & _
                         " WHERE Odate = '" & m_typDetail_Rec(intIndex1).Odate & "'" & _
                         " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode & _
                         " ORDER BY Bcode,Odate,Num"
                adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
                Do While Not adoRecordset1.EOF
                    adoRecordset1.Fields("Rdiv") = PAYMENT_ON
                    adoRecordset1.Fields("R") = intR
                    adoRecordset1.Fields("Rdate") = strRdate
                    adoRecordset1.Update
                
                    adoRecordset1.MoveNext
                Loop
                adoRecordset1.Close
            Else
                m_typDetail_Rec(intIndex1).Zandaka = m_typDetail_Rec(intIndex1).Zandaka - m_typDetail_Rec(intIndex1).Nyukin
                m_typDetail_Rec(intIndex1).KeepZan = 0
                m_typDetail_Rec(intIndex1).Nyukin = 0
                m_typDetail_Rec(intIndex1).Keep = 0
                '201107
                m_typDetail_Rec(intIndex1).Brate2Zan = 0
            End If
        End If
    Next intIndex1
            
    g_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
            
    Call Detail_Dislplay(1)
    Call Detail_ScrollBar
    Call Calc_Total

'    cboInpR.Text = ""
    imnInpRdate_Year.Value = Year(Now())
'    imnInp_Rdate_Month.Value = Month(Now())
    imnInp_Rdate_Day.Text = ""
    imnInp_Nyukin.Value = 0
    fraInput.Enabled = False
            
    cboBcode(0).SetFocus
            
    Exit Sub

cmdInput_Click_Err:

    Screen.MousePointer = vbDefault
    g_clsAdoSQL.Connection.RollbackTrans

    Call MsgBox("入金クリック時エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdInput_Click_Err")
            
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

    Call MsgBox("開催年月日と担当者の変更クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdLogin_Click_Err")

End Sub

'目　的　　：
'条　件　　：入金ボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub cmdPayment_Click(Index As Integer)

    On Error GoTo cmdPayment_Click_Err
    
    If cmdPayment(Index).Caption = CMD_PAYMENT_OFF Then
        If MsgBox("入金処理しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
        '省略された場合の初期値をセット
        If imnNyukin(Index).Value = 0 Then
            imnNyukin(Index).Value = imnZandaka(Index).Value
        End If
        If Trim(cboR(Index).Text) = "" Then cboR(Index).Text = PAYMENT_DIV1_NAME
        If Trim(imnRdate_Year(Index).Text) = "" And _
           Trim(imnRdate_Month(Index).Text) = "" And _
           Trim(imnRdate_Day(Index).Text) = "" Then
            imnRdate_Year(Index).Value = Format(Now(), "yyyy")
            imnRdate_Month(Index).Value = Format(Now(), "m")
            imnRdate_Day(Index).Value = Format(Now(), "d")
        End If
        DoEvents
    
        '入力チェック
        If Detail_DoValidationChecks(Index) = False Then Exit Sub
    
        '入金処理
        If PayMent_Data(Index, True) = False Then Exit Sub
        
        '残高がゼロになった場合のみ
        If (CCur(imnZandaka(Index).Value) - CCur(imnNyukin(Index).Value)) <= 0 Then
            fraDetail(Index).BackColor = BACK_COLOR_ON
            cmdPayment(Index).Caption = CMD_PAYMENT_ON
            cboR(Index).Enabled = False
            imnRdate_Year(Index).Enabled = False
            imnRdate_Month(Index).Enabled = False
            imnRdate_Day(Index).Enabled = False
        Else
            imnNyukin(Index).Value = 0
            imnKeep(Index).Value = 0
            cboR(Index).Text = ""
            imnRdate_Year(Index).Text = ""
            imnRdate_Month(Index).Text = ""
            imnRdate_Day(Index).Text = ""
            '201107
            imnBrate2(Index).Value = 0
        End If
    
    ElseIf cmdPayment(Index).Caption = CMD_PAYMENT_ON Then
        If MsgBox("入金を取り消しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub
    
        cboR(Index).Text = ""
        imnRdate_Year(Index).Text = ""
        imnRdate_Month(Index).Text = ""
        imnRdate_Day(Index).Text = ""
            
        '入金取り消し
        If PayMent_Data(Index, False) = False Then Exit Sub
        
        fraDetail(Index).BackColor = BACK_COLOR_OFF
        cmdPayment(Index).Caption = CMD_PAYMENT_OFF
        cboR(Index).Enabled = True
        imnRdate_Year(Index).Enabled = True
        imnRdate_Month(Index).Enabled = True
        imnRdate_Day(Index).Enabled = True
    End If
    
    '小計などの計算
    Call Calc_Total
    DoEvents

    Exit Sub

cmdPayment_Click_Err:

    Call MsgBox("入金ボタンクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdPayment_Click_Err")

End Sub

Private Sub cmdReset_Click()
    
    If cboBcode(0).Text = "" Then
        Call MsgBox("買主コードを入力してください", vbOKOnly + vbInformation, "")
        cboBcode(0).SetFocus
        DoEvents
        Exit Sub
    End If
    
    frmTorikesi.g_intBocode = cboBcode(0).Text
    frmTorikesi.Show vbModal
    
    '取消してたら検索し直し
    If frmTorikesi.g_blnTorikesizumi = True Then
        Call cmdSearch_Click
    End If

End Sub

Private Sub cmdSearch_Click()

    Dim intIndex1 As Integer

    On Error GoTo cmdSearch_Click_Err

    If cboBcode(0).Text = "" Then
        Call MsgBox("買主コードを入力してください", vbOKOnly + vbInformation, "")
        cboBcode(0).SetFocus
        DoEvents
        Exit Sub
    End If

    '明細クリア
    For intIndex1 = 1 To DETAIL_MAX
        Call Detail_Clear(intIndex1)
    Next intIndex1
    imnGtotal_Total.Value = 0
    
    Erase m_typDetail_Rec   '配列初期化
    
    If optDisplayFlg(0).Value = True Then
        Call Detail_SetData(0)
    ElseIf optDisplayFlg(1).Value = True Then
        Call Detail_SetData(1)
    End If
    Call Detail_Dislplay(1)
    Call Detail_ScrollBar
    
    If UBound(m_typDetail_Rec) <= 0 Then
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
        cboBcode(0).SetFocus
        DoEvents
        Exit Sub
    End If

    fraInput.Enabled = True
    cboInpR.SetFocus
    DoEvents

    Exit Sub
    
cmdSearch_Click_Err:

    Call MsgBox("検索開始クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdSearch_Click_Err")

End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Form_KeyDown_Err
    
    'リターンキーで次のコントロールへフォーカス移動
    If KeyCode = vbKeyReturn And Shift <> vbShiftMask Then
        KeyCode = 0
        Global_SendKeys Me, VK_TAB
        Exit Sub
    End If
    
    'ショートカットキーの割り当て
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
            cmdExecute.SetFocus
            DoEvents
            Call cmdExecute_Click
        Case vbKeyF11
        Case vbKeyF12
            cmdInput.SetFocus
            DoEvents
            Call cmdInput_Click
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

    Call MsgBox("フォームキーダウン時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")

End Sub

'目　的　　：
'条　件　　：フォームロード時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub Form_Load()

    Dim intIndex1 As Integer

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "入金入力"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    optDisplayFlg(0).Value = True
    
    cboInpR.Clear
    cboInpR.AddItem PAYMENT_DIV1_NAME
    cboInpR.AddItem PAYMENT_DIV2_NAME
    cboInpR.AddItem PAYMENT_DIV3_NAME
    
    'コンボボックス作成
    For intIndex1 = 1 To DETAIL_MAX
        cboR(intIndex1).Clear
        cboR(intIndex1).AddItem PAYMENT_DIV1_NAME
        cboR(intIndex1).AddItem PAYMENT_DIV2_NAME
        cboR(intIndex1).AddItem PAYMENT_DIV3_NAME
    Next intIndex1
    
    Call FieldsClear(0)
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    End

End Sub

'目　的　　：
'条　件　　：フォームアンロード時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set g_clsAdoSQL = Nothing
    Set g_clsReg = Nothing
    End
    
    Exit Sub
    
Form_Unload_Err:

    Call MsgBox("フォームアンロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")
    End

End Sub

'目　的　　：画面クリア
'条　件　　：
'結　果　　：
'引　数　　：0：全画面
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    Dim intIndex1 As Integer

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        cboBcode(0).Text = ""
        cboBcode(1).Text = ""
        lblScode_Name(0).Caption = ""
        lblScode_Name(1).Caption = ""
        
        cboInpR.Text = ""
        imnInpRdate_Year.Value = Year(Now())
        imnInp_Rdate_Month.Value = Month(Now())
        imnInp_Rdate_Day.Text = ""
        imnInp_Nyukin.Value = 0
        fraInput.Enabled = False
        
        '明細クリア
        For intIndex1 = 1 To DETAIL_MAX
            Call Detail_Clear(intIndex1)
        Next intIndex1
        imnGtotal_Total.Value = 0
        imnZandaka_Total.Value = 0
        imnKeep_Total.Value = 0
        '201107
        imnBrate2_Total.Value = 0
        
        Erase m_typDetail_Rec   '配列初期化
        ReDim m_typDetail_Rec(0)
        
'        If optDisplayFlg(0).Value = True Then
'            Call Detail_SetData(0)
'        ElseIf optDisplayFlg(1).Value = True Then
'            Call Detail_SetData(1)
'        End If

        Call Detail_Dislplay(1)
        Call Detail_ScrollBar
        
'        If UBound(m_typDetail_Rec) <= 0 Then
'            Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
'        End If
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "開催年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function

Private Sub imnInp_Nyukin_GotFocus()

    imnInp_Nyukin.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnInp_Nyukin_LostFocus()

    imnInp_Nyukin.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnInp_Rdate_Day_GotFocus()

    imnInp_Rdate_Day.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnInp_Rdate_Day_LostFocus()

    imnInp_Rdate_Day.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnInp_Rdate_Month_GotFocus()

    imnInp_Rdate_Month.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnInp_Rdate_Month_LostFocus()

    imnInp_Rdate_Month.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnInpRdate_Year_GotFocus()

    imnInpRdate_Year.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnInpRdate_Year_LostFocus()

    imnInpRdate_Year.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnKeep_GotFocus(Index As Integer)

    imnKeep(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnKeep_LostFocus(Index As Integer)

    imnKeep(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnBrate2_GotFocus(Index As Integer)

    imnBrate2(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBrate2_LostFocus(Index As Integer)

    imnBrate2(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnNyukin_GotFocus(Index As Integer)

    imnNyukin(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnNyukin_LostFocus(Index As Integer)

    imnNyukin(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRdate_Day_GotFocus(Index As Integer)

    imnRdate_Day(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Day_LostFocus(Index As Integer)

    imnRdate_Day(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRdate_Month_GotFocus(Index As Integer)

    imnRdate_Month(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Month_LostFocus(Index As Integer)

    imnRdate_Month(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRdate_Year_GotFocus(Index As Integer)

    imnRdate_Year(Index).BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnRdate_Year_LostFocus(Index As Integer)

    imnRdate_Year(Index).BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cboBcode(0).SetFocus

End Sub

'目　的　　：明細ワークへデータセット
'条　件　　：
'結　果　　：
'引　数　　：intFlg 0:未精算データのみ 1:全データ
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Function Detail_SetData(intFlg As Integer) As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim intIndex1 As Integer
    Dim intRecordCount As Integer
    Dim strBuff As String
    Dim strBcodeFrom As String
    Dim strBcodeTo As String
    Dim typDetail_Sort() As Detail_Record

    On Error GoTo Detail_SetData_Err
    
    Detail_SetData = False
    
    Screen.MousePointer = vbHourglass
    
    '買主コード抽出
    If Trim(cboBcode(0).Text) <> "" Then
        strBcodeFrom = cboBcode(0).Text
    Else
        strBcodeFrom = "1"
    End If
    If Trim(cboBcode(1).Text) <> "" Then
        strBcodeTo = cboBcode(1).Text
    Else
        strBcodeTo = "9999"
    End If
    
    '初期化
    ReDim typDetail_Sort(0)
    
'********** 買主精算データ **********
    
    If intFlg = 0 Then
        strSQL = "{call sp_YPMF0601;1('" & lblOdate.Caption & "'," & strBcodeFrom & "," & strBcodeTo & ")}"
    Else
        strSQL = "{call sp_YPMF0601;2('" & lblOdate.Caption & "'," & strBcodeFrom & "," & strBcodeTo & ")}"
    End If
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        intRecordCount = adoRecordset1.RecordCount
        ReDim typDetail_Sort(intRecordCount)
    
        For intIndex1 = 1 To intRecordCount
            typDetail_Sort(intIndex1).Bcode = IIf(IsNull(adoRecordset1.Fields("Bcode")), "", adoRecordset1.Fields("Bcode"))
            typDetail_Sort(intIndex1).Bname = Global_Get_Bname(g_clsAdoSQL, typDetail_Sort(intIndex1).Bcode, lblOdate.Caption, strBuff)
            typDetail_Sort(intIndex1).Odate = IIf(IsNull(adoRecordset1.Fields("Odate")), "", adoRecordset1.Fields("Odate"))
            typDetail_Sort(intIndex1).Gtotal = IIf(IsNull(adoRecordset1.Fields("Gtotal")), 0, adoRecordset1.Fields("Gtotal"))
            typDetail_Sort(intIndex1).KeepZan = IIf(IsNull(adoRecordset1.Fields("Keep")), 0, adoRecordset1.Fields("Keep"))
            '201107
            typDetail_Sort(intIndex1).Brate2Zan = IIf(IsNull(adoRecordset1.Fields("Brate2")), 0, adoRecordset1.Fields("Brate2"))
            
            If intFlg = 0 Then
                typDetail_Sort(intIndex1).Rdiv = PAYMENT_OFF
                typDetail_Sort(intIndex1).R = 0
                typDetail_Sort(intIndex1).Rdate = ""
            ElseIf intFlg = 1 Then
                typDetail_Sort(intIndex1).Rdiv = PAYMENT_OFF
                typDetail_Sort(intIndex1).R = 0
                typDetail_Sort(intIndex1).Rdate = ""
                '入金区分
                If Not IsNull(adoRecordset1.Fields("Rdiv")) Then
                    If adoRecordset1.Fields("Rdiv") = PAYMENT_ON Then
                        typDetail_Sort(intIndex1).Rdiv = PAYMENT_ON
                        typDetail_Sort(intIndex1).R = IIf(IsNull(adoRecordset1.Fields("R")), 0, adoRecordset1.Fields("R"))
                        typDetail_Sort(intIndex1).Rdate = IIf(IsNull(adoRecordset1.Fields("Rdate")), "", adoRecordset1.Fields("Rdate"))
                    End If
                End If
            End If
            
            adoRecordset1.MoveNext
        Next intIndex1
    End If
    adoRecordset1.Close
    
'********** 買主精算累積データ **********
    
'    If intFlg = 0 Then
'        strSQL = "{call sp_YPMF0602;1('" & lblOdate.Caption & "'," & strBcodeFrom & "," & strBcodeTo & ")}"
'    Else
'        strSQL = "{call sp_YPMF0602;2('" & lblOdate.Caption & "'," & strBcodeFrom & "," & strBcodeTo & ")}"
'    End If
'    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
'    If adoRecordset1.EOF = False Then
'        intRecordCount = adoRecordset1.RecordCount
'        ReDim Preserve typDetail_Sort(UBound(typDetail_Sort) + intRecordCount)
'
'        For intIndex1 = 1 To intRecordCount
'            typDetail_Sort(intIndex1).Bcode = IIf(IsNull(adoRecordset1.Fields("Bcode")), "", adoRecordset1.Fields("Bcode"))
'            typDetail_Sort(intIndex1).Bname = Global_Get_Bname(g_clsAdoSQL, typDetail_Sort(intIndex1).Bcode, lblOdate.Caption, strBuff)
'            typDetail_Sort(intIndex1).Odate = IIf(IsNull(adoRecordset1.Fields("Odate")), "", adoRecordset1.Fields("Odate"))
'            typDetail_Sort(intIndex1).Gtotal = IIf(IsNull(adoRecordset1.Fields("Gtotal")), 0, adoRecordset1.Fields("Gtotal"))
'            If intFlg = 0 Then
'                typDetail_Sort(intIndex1).Rdiv = PAYMENT_OFF
'                typDetail_Sort(intIndex1).R = 0
'                typDetail_Sort(intIndex1).Rdate = ""
'            ElseIf intFlg = 1 Then
'                typDetail_Sort(intIndex1).Rdiv = PAYMENT_OFF
'                typDetail_Sort(intIndex1).R = 0
'                typDetail_Sort(intIndex1).Rdate = ""
'                '入金区分
'                If Not IsNull(adoRecordset1.Fields("Rdiv")) Then
'                    If adoRecordset1.Fields("Rdiv") = PAYMENT_ON Then
'                        typDetail_Sort(intIndex1).Rdiv = PAYMENT_ON
'                        typDetail_Sort(intIndex1).R = IIf(IsNull(adoRecordset1.Fields("R")), 0, adoRecordset1.Fields("R"))
'                        typDetail_Sort(intIndex1).Rdate = IIf(IsNull(adoRecordset1.Fields("Rdate")), "", adoRecordset1.Fields("Rdate"))
'                    End If
'                End If
'            End If
'
'            adoRecordset1.MoveNext
'        Next intIndex1
'    End If
'    adoRecordset1.Close
    
    '買主コード、開催日でソート
    Call Detail_Sort(typDetail_Sort, m_typDetail_Rec)
    
    '残高の取得
    Call Get_Zandaka

    '小計などの計算
    Call Calc_Total
    
    Screen.MousePointer = vbDefault
    
    Detail_SetData = True
    
    Exit Function

Detail_SetData_Err:

    Detail_SetData = False
    Screen.MousePointer = vbDefault
    Call MsgBox("フィールドセットエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_SetData_Err")

End Function

'目　的　　：合計の計算
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub Calc_Total()

    Dim curBuff(4) As Currency
    Dim intIndex1 As Integer

    On Error GoTo Calc_Total_Err
    
    curBuff(1) = 0: curBuff(2) = 0: curBuff(3) = 0: curBuff(4) = 0
    
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
         curBuff(1) = curBuff(1) + m_typDetail_Rec(intIndex1).Gtotal
         curBuff(2) = curBuff(2) + m_typDetail_Rec(intIndex1).Zandaka
         curBuff(3) = curBuff(3) + m_typDetail_Rec(intIndex1).Keep
         '201107
         curBuff(4) = curBuff(4) + m_typDetail_Rec(intIndex1).Brate2
    Next intIndex1
    
    imnGtotal_Total.Value = curBuff(1)
    imnZandaka_Total.Value = curBuff(2)
    imnKeep_Total.Value = curBuff(3)
    '201107
    imnBrate2_Total.Value = curBuff(4)
    
    Exit Sub
    
Calc_Total_Err:

    Call MsgBox("合計の計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Total_Err")

End Sub

'目　的　　：明細の表示
'条　件　　：
'結　果　　：
'引　数　　：(intStartLine：表示開始行番号)
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Function Detail_Dislplay(intStartLine As Integer) As Boolean

    Dim intIndex1 As Integer
    Dim intPostion As Integer

    On Error GoTo Detail_Dislplay_Err
    
    Detail_Dislplay = False
    
    intPostion = intStartLine
    For intIndex1 = 1 To DETAIL_MAX
        '明細の１行クリア
        Call Detail_Clear(intIndex1)
        
        If intPostion <= UBound(m_typDetail_Rec) Then
            fraDetail(intIndex1).Visible = True
        
            lblBcode(intIndex1).Caption = m_typDetail_Rec(intPostion).Bcode
            lblBname(intIndex1).Caption = m_typDetail_Rec(intPostion).Bname
            lblOdate_Detail(intIndex1).Caption = m_typDetail_Rec(intPostion).Odate
            imnGtotal(intIndex1).Value = m_typDetail_Rec(intPostion).Gtotal
            imnZandaka(intIndex1).Value = m_typDetail_Rec(intPostion).Zandaka
            imnKeepZan(intIndex1).Value = m_typDetail_Rec(intPostion).KeepZan
            imnNyukin(intIndex1).Value = m_typDetail_Rec(intPostion).Nyukin
            imnKeep(intIndex1).Value = m_typDetail_Rec(intPostion).Keep
            '201107
            imnBrate2Zan(intIndex1).Value = m_typDetail_Rec(intPostion).Brate2Zan
            imnBrate2(intIndex1).Value = m_typDetail_Rec(intPostion).Brate2
            
            '入金種別
            Select Case m_typDetail_Rec(intPostion).R
                Case PAYMENT_DIV_CASH:
                    cboR(intIndex1).Text = PAYMENT_DIV1_NAME
                Case SHIHARAI_DIV_CHECK:
                    cboR(intIndex1).Text = PAYMENT_DIV2_NAME
                Case SHIHARAI_DIV_TRANSFER:
                    cboR(intIndex1).Text = PAYMENT_DIV3_NAME
                Case Else
                    cboR(intIndex1).Text = ""
            End Select
            '入金日
            If m_typDetail_Rec(intPostion).Rdate <> "" Then
                imnRdate_Year(intIndex1).Text = left$(m_typDetail_Rec(intPostion).Rdate, 4)
                imnRdate_Month(intIndex1).Text = Mid$(m_typDetail_Rec(intPostion).Rdate, 6, 2)
                imnRdate_Day(intIndex1).Text = right(m_typDetail_Rec(intPostion).Rdate, 2)
            Else
                imnRdate_Year(intIndex1).Text = ""
                imnRdate_Month(intIndex1).Text = ""
                imnRdate_Day(intIndex1).Text = ""
            End If
            '入金区分
            If m_typDetail_Rec(intPostion).Rdiv = PAYMENT_OFF Then
                cmdPayment(intIndex1).Caption = CMD_PAYMENT_OFF
            ElseIf m_typDetail_Rec(intPostion).Rdiv = PAYMENT_ON Then
                cmdPayment(intIndex1).Caption = CMD_PAYMENT_ON
            End If
            '背景色
            If m_typDetail_Rec(intPostion).Rdiv = PAYMENT_ON Then
                fraDetail(intIndex1).BackColor = BACK_COLOR_ON
                cboR(intIndex1).Enabled = False
                imnRdate_Year(intIndex1).Enabled = False
                imnRdate_Month(intIndex1).Enabled = False
                imnRdate_Day(intIndex1).Enabled = False
            Else
                fraDetail(intIndex1).BackColor = BACK_COLOR_OFF
                cboR(intIndex1).Enabled = True
                imnRdate_Year(intIndex1).Enabled = True
                imnRdate_Month(intIndex1).Enabled = True
                imnRdate_Day(intIndex1).Enabled = True
            End If
        End If
        intPostion = intPostion + 1
    Next intIndex1
    
    Detail_Dislplay = True
    
    Exit Function
    
Detail_Dislplay_Err:

    Detail_Dislplay = False
    Call MsgBox("明細の表示エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Dislplay_Err")

End Function

'目　的　　：明細の１行クリア
'条　件　　：
'結　果　　：
'引　数　　：(intClearLine:絶対位置)
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Function Detail_Clear(intClearLine As Integer) As Boolean

    On Error GoTo Detail_Clear_Err
    
    lblBcode(intClearLine).Caption = ""
    lblBname(intClearLine).Caption = ""
    lblOdate_Detail(intClearLine).Caption = ""
    imnGtotal(intClearLine).Value = 0
    imnZandaka(intClearLine).Value = 0
    imnKeepZan(intClearLine).Value = 0
    imnNyukin(intClearLine).Value = 0
    cboR(intClearLine).Text = ""
    imnRdate_Year(intClearLine).Text = ""
    imnRdate_Month(intClearLine).Text = ""
    imnRdate_Day(intClearLine).Text = ""
    cmdPayment(intClearLine).Caption = ""
    imnKeep(intClearLine).Value = 0
    '201107
    imnBrate2Zan(intClearLine).Value = 0
    imnBrate2(intClearLine).Value = 0
    
    fraDetail(intClearLine).BackColor = BACK_COLOR_OFF
    fraDetail(intClearLine).Visible = False
    
    Detail_Clear = True
    
    Exit Function
    
Detail_Clear_Err:

    Detail_Clear = False
    Call MsgBox("明細の１行クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Clear_Err")

End Function

'目　的　　：スクロールバーの制御
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
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
    Call MsgBox("スクロールバーの制御エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_ScrollBar_Err")

End Function

Private Sub VScroll1_Change()

    If VScroll1.Tag = "EventFalse" Then Exit Sub

    Call Detail_Dislplay(VScroll1.Value)

End Sub

'目　的　　：入金登録
'条　件　　：
'結　果　　：
'引　数　　：Index:配列位置 blnFlg(True:入金 False:入金取り消し)
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Function PayMent_Data(Index As Integer, blnFlg As Boolean) As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strRdate As String
    Dim intR As Integer
    
    On Error GoTo PayMent_Data_Err
    
    Screen.MousePointer = vbHourglass
    
    PayMent_Data = False
    
    '入金種別
    Select Case Trim(cboR(Index).Text)
        Case PAYMENT_DIV1_NAME:
            intR = PAYMENT_DIV_CASH
        Case PAYMENT_DIV2_NAME:
            intR = PAYMENT_DIV_CHECK
        Case PAYMENT_DIV3_NAME:
            intR = PAYMENT_DIV_TRANSFER
        Case Else
            intR = PAYMENT_DIV_CASH
    End Select
    '入金日
    strRdate = Format(imnRdate_Year(Index).Text, "0000") & "/" & Format(imnRdate_Month(Index).Text, "00") & "/" & Format(imnRdate_Day(Index).Text, "00")
    
    g_clsAdoSQL.Connection.BeginTrans
        
    If blnFlg = True Then
        '********** 入金処理 **********
    
        '入金データ
        strSQL = "SELECT * FROM DT060" & _
                 " WHERE Odate = '" & lblOdate_Detail(Index).Caption & "'" & _
                 " AND Bcode = " & lblBcode(Index).Caption & _
                 " AND Rdate = '" & strRdate & "'" & _
                 " ORDER BY Odate,Bcode,Rdate"
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If adoRecordset1.EOF = True Then
            adoRecordset1.AddNew
            adoRecordset1.Fields("Odate") = lblOdate_Detail(Index).Caption
            adoRecordset1.Fields("Bcode") = lblBcode(Index).Caption
            adoRecordset1.Fields("Rdate") = strRdate
            adoRecordset1.Fields("Rdiv") = PAYMENT_ON
            adoRecordset1.Fields("R") = intR
            '201107
            'adoRecordset1.Fields("Ptotal") = CCur(imnNyukin(Index).Value) + CCur(imnKeep(Index).Value)
            adoRecordset1.Fields("Ptotal") = CCur(imnNyukin(Index).Value) + CCur(imnKeep(Index).Value) + CCur(imnBrate2(Index).Value)
            adoRecordset1.Fields("Ptotal2") = CCur(imnKeep(Index).Value)
            '201107
            adoRecordset1.Fields("Ptotal3") = CCur(imnBrate2(Index).Value)
            adoRecordset1.Update
        Else
            '入金日が同じデータがあった場合、金額を合計する
            adoRecordset1.Fields("Rdiv") = PAYMENT_ON
            adoRecordset1.Fields("R") = intR
            '201107
            'adoRecordset1.Fields("Ptotal") = CCur(adoRecordset1.Fields("Ptotal")) + CCur(imnNyukin(Index).Value) + CCur(imnKeep(Index).Value)
            adoRecordset1.Fields("Ptotal") = CCur(adoRecordset1.Fields("Ptotal")) + CCur(imnNyukin(Index).Value) + CCur(imnKeep(Index).Value) + CCur(imnBrate2(Index).Value)
            adoRecordset1.Fields("Ptotal2") = CCur(adoRecordset1.Fields("Ptotal2")) + CCur(imnKeep(Index).Value)
            '201107
            adoRecordset1.Fields("Ptotal3") = CCur(adoRecordset1.Fields("Ptotal3")) + CCur(imnBrate2(Index).Value)
            adoRecordset1.Update
        End If
        adoRecordset1.Close
    
        '残高がゼロになった場合のみ 201107
'        If (CCur(imnZandaka(Index).Value) - CCur(imnNyukin(Index).Value) - CCur(imnKeep(Index).Value)) <= 0 Then
        If (CCur(imnZandaka(Index).Value) - CCur(imnNyukin(Index).Value) - CCur(imnKeep(Index).Value) - CCur(imnBrate2(Index).Value)) <= 0 Then
            imnZandaka(Index).Value = 0
            imnNyukin(Index).Value = 0
            imnKeep(Index).Value = 0
            imnBrate2(Index).Value = 0
            
            '買主精算データ
            strSQL = "SELECT * FROM DT041" & _
                     " WHERE Odate = '" & lblOdate_Detail(Index).Caption & "'" & _
                     " AND Bcode = " & lblBcode(Index).Caption & _
                     " ORDER BY Bcode,Odate,Num"
            adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            Do While Not adoRecordset1.EOF
                adoRecordset1.Fields("Rdiv") = PAYMENT_ON
                adoRecordset1.Fields("R") = intR
                adoRecordset1.Fields("Rdate") = strRdate
                adoRecordset1.Update
            
                adoRecordset1.MoveNext
            Loop
            adoRecordset1.Close
        
            '買主精算データ(累積)
            strSQL = "SELECT * FROM RT041" & _
                     " WHERE Odate = '" & lblOdate_Detail(Index).Caption & "'" & _
                     " AND Bcode = " & lblBcode(Index).Caption & _
                     " ORDER BY Bcode,Odate,Num"
            adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            Do While Not adoRecordset1.EOF
                adoRecordset1.Fields("Rdiv") = PAYMENT_ON
                adoRecordset1.Fields("R") = intR
                adoRecordset1.Fields("Rdate") = strRdate
                adoRecordset1.Update
            
                adoRecordset1.MoveNext
            Loop
            adoRecordset1.Close
        Else
            '201107
            'imnZandaka(Index).Value = CCur(imnZandaka(Index).Value) - CCur(imnNyukin(Index).Value) - CCur(imnKeep(Index).Value)
            imnZandaka(Index).Value = CCur(imnZandaka(Index).Value) - CCur(imnNyukin(Index).Value) - CCur(imnKeep(Index).Value) - CCur(imnBrate2(Index).Value)
            imnNyukin(Index).Value = 0
            imnKeep(Index).Value = 0
            imnBrate2(Index).Value = 0
        End If
    Else
        '********** 入金取り消し処理 **********
                
        '入金データ削除
        strSQL = "DELETE FROM DT060" & _
                 " WHERE Odate = '" & lblOdate_Detail(Index).Caption & "'" & _
                 " AND Bcode = " & lblBcode(Index).Caption
        g_clsAdoSQL.Connection.Execute strSQL
        
        
        '買主精算データ
        strSQL = "SELECT * FROM DT041" & _
                 " WHERE Odate = '" & lblOdate_Detail(Index).Caption & "'" & _
                 " AND Bcode = " & lblBcode(Index).Caption & _
                 " ORDER BY Bcode,Odate,Num"
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        Do While Not adoRecordset1.EOF
            adoRecordset1.Fields("Rdiv") = PAYMENT_OFF
            adoRecordset1.Fields("R") = 0
            adoRecordset1.Fields("Rdate") = ""
            adoRecordset1.Update
        
            adoRecordset1.MoveNext
        Loop
        adoRecordset1.Close
        
        '買主精算データ(累積)
        strSQL = "SELECT * FROM RT041" & _
                 " WHERE Odate = '" & lblOdate_Detail(Index).Caption & "'" & _
                 " AND Bcode = " & lblBcode(Index).Caption & _
                 " ORDER BY Bcode,Odate,Num"
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        Do While Not adoRecordset1.EOF
            adoRecordset1.Fields("Rdiv") = PAYMENT_OFF
            adoRecordset1.Fields("R") = 0
            adoRecordset1.Fields("Rdate") = ""
            adoRecordset1.Update
        
            adoRecordset1.MoveNext
        Loop
        adoRecordset1.Close
        
        '残高金額
        imnZandaka(Index).Value = imnGtotal(Index).Value
    End If
        
    g_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    PayMent_Data = True
    
    Exit Function
    
PayMent_Data_Err:

    PayMent_Data = False
    Screen.MousePointer = vbDefault
    g_clsAdoSQL.Connection.RollbackTrans
    Call MsgBox("入金登録エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "PayMent_Data_Err")

End Function

'目　的　　：明細の入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Function Detail_DoValidationChecks(Index As Integer) As Boolean

    Dim strErrMsg As String
        
    On Error GoTo Detail_DoValidationChecks_Err

    If Trim(lblBcode(Index).Caption) = "" Then
        strErrMsg = "買主コードが入力されていません。"
        GoTo ErrorTrap:
    End If
    If Trim(lblOdate_Detail(Index).Caption) = "" Then
        strErrMsg = "開催日が入力されていません。"
        GoTo ErrorTrap:
    End If
    If Trim(cboR(Index).Text) = "" Then
        strErrMsg = "開催日が入力されていません。"
        cboR(Index).SetFocus
        GoTo ErrorTrap:
    End If
    If imnNyukin(Index).Value = 0 Then
        strErrMsg = "今回入金額が入力されていません。"
        imnNyukin(Index).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Year(Index).Text) = "" Then
        strErrMsg = "入金日が入力されていません。"
        imnRdate_Year(Index).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Month(Index).Text) = "" Then
        strErrMsg = "入金日が入力されていません。"
        imnRdate_Month(Index).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Day(Index).Text) = "" Then
        strErrMsg = "入金日が入力されていません。"
        imnRdate_Day(Index).SetFocus
        GoTo ErrorTrap:
    End If
    If Global_IsDate(imnRdate_Year(Index).Text, imnRdate_Month(Index).Text, imnRdate_Day(Index).Text) = False Then
        strErrMsg = "正しい入金日が入力されていません。"
        imnRdate_Day(Index).SetFocus
        GoTo ErrorTrap:
    End If
    
    Detail_DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    Detail_DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
Detail_DoValidationChecks_Err:

    Detail_DoValidationChecks = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_DoValidationChecks_Err")

End Function

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
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
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboBcode_Err")

End Sub

'目　的　　：明細のソート
'条　件　　：買主コードの昇順
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／０８
'更新履歴　：
'
Private Function Detail_Sort(ByRef Before() As Detail_Record, ByRef After() As Detail_Record) As Boolean

    Dim intIndex1 As Integer
    Dim intPostion As Integer
    Dim work(1) As Detail_Record     'ワーク

    On Error GoTo Detail_Sort_Err
    
    If UBound(Before) <= 0 Then
        ReDim After(0)
        Exit Function
    End If
    
    'バブルソート
    For intIndex1 = UBound(Before) To 1 Step -1
        For intPostion = 1 To intIndex1 - 1
            If CInt(Before(intPostion).Bcode) >= CInt(Before(intPostion + 1).Bcode) And Before(intPostion).Odate >= Before(intPostion + 1).Odate Then
                work(1) = Before(intPostion)
                Before(intPostion) = Before(intPostion + 1)
                Before(intPostion + 1) = work(1)
            End If
        Next intPostion
    Next intIndex1
    
    '配列コピー
    ReDim After(UBound(Before))
    After = Before
    
    Detail_Sort = True
    
    Exit Function
    
Detail_Sort_Err:

    Detail_Sort = False
    Call MsgBox("明細のソートエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Sort_Err")

End Function

'目　的　　：残高の取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２０
'更新履歴　：
'
Private Sub Get_Zandaka()

    Dim intIndex1 As Integer
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo Get_Zandaka_Err
    
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
        '入金区分
        If m_typDetail_Rec(intIndex1).Rdiv = PAYMENT_ON Then
            m_typDetail_Rec(intIndex1).Zandaka = 0
        Else
            m_typDetail_Rec(intIndex1).Zandaka = m_typDetail_Rec(intIndex1).Gtotal
            
            '入金データ
            strSQL = "SELECT * FROM DT060" & _
                     " WHERE Odate = '" & m_typDetail_Rec(intIndex1).Odate & "'" & _
                     " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode
            adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            Do While Not adoRecordset1.EOF
                If Not IsNull(adoRecordset1.Fields("Ptotal")) Then
                    '残高金額＝残高金額－入金金額合計
                    m_typDetail_Rec(intIndex1).Zandaka = m_typDetail_Rec(intIndex1).Zandaka - CCur(adoRecordset1.Fields("Ptotal"))
                    
                    '残高がゼロ以下の場合は、買主精算データを更新
                    If m_typDetail_Rec(intIndex1).Zandaka <= 0 Then
                        strSQL = "UPDATE DT041"
                        strSQL = strSQL & " SET Rdiv = " & PAYMENT_ON & ","
                        strSQL = strSQL & " R = " & adoRecordset1.Fields("R") & ","
                        strSQL = strSQL & " Rdate = '" & adoRecordset1.Fields("Rdate") & "'"
                        strSQL = strSQL & " WHERE Odate = '" & m_typDetail_Rec(intIndex1).Odate & "'"
                        strSQL = strSQL & " AND Bcode = " & m_typDetail_Rec(intIndex1).Bcode
                        g_clsAdoSQL.Connection.Execute strSQL
                        Exit Do
                    End If
                End If
                adoRecordset1.MoveNext
            Loop
            adoRecordset1.Close
        End If
    Next intIndex1
    
    Exit Sub
    
Get_Zandaka_Err:

    Call MsgBox("残高の取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Get_Zandaka_Err")

End Sub

