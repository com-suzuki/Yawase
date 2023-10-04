VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmYpmf110 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   10425
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12330
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf110.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   12330
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   33
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton cmdLogin 
         Caption         =   "開催年月日と担当者の変更"
         Height          =   375
         Left            =   6960
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9660
         TabIndex        =   34
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
         Contents        =   "frmYpmf110.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   32
      Top             =   9600
      Width           =   12135
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "画面ｸﾘｱ(F8)"
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
         rText.left      =   10
         rText.top       =   8
         rText.right     =   103
         rText.bottom    =   27
         Picture         =   "frmYpmf110.frx":0D13
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10320
         TabIndex        =   28
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
         Picture         =   "frmYpmf110.frx":0D2F
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   8580
         TabIndex        =   27
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "印刷(F12)"
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
         rPic.left       =   10
         rPic.top        =   8
         rPic.right      =   16
         rPic.bottom     =   16
         rText.left      =   30
         rText.top       =   8
         rText.right     =   105
         rText.bottom    =   27
         Picture         =   "frmYpmf110.frx":0E89
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdCalc 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "再計算(F7)"
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
         rText.left      =   15
         rText.top       =   8
         rText.right     =   97
         rText.bottom    =   27
         Picture         =   "frmYpmf110.frx":0F9B
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   8835
      Left            =   60
      TabIndex        =   31
      Top             =   660
      Width           =   12135
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   180
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "現　金"
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
         LabelWidth      =   40
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
      Begin imNumber6Ctl.imNumber imnCashe 
         Height          =   435
         Left            =   1920
         TabIndex        =   1
         Top             =   180
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":0FB7
         Caption         =   "frmYpmf110.frx":0FD7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1045
         Keys            =   "frmYpmf110.frx":1063
         Spin            =   "frmYpmf110.frx":10AD
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
         ValueVT         =   2012217349
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   660
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "利　益"
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
         LabelWidth      =   40
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
      Begin imNumber6Ctl.imNumber imnProfit 
         Height          =   435
         Left            =   1920
         TabIndex        =   2
         Top             =   660
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":10D5
         Caption         =   "frmYpmf110.frx":10F5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1163
         Keys            =   "frmYpmf110.frx":1181
         Spin            =   "frmYpmf110.frx":11CB
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
         ValueVT         =   5
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   1140
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "未収分の入金"
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
         LabelWidth      =   90
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
      Begin imNumber6Ctl.imNumber imnRtotal 
         Height          =   435
         Left            =   1920
         TabIndex        =   3
         Top             =   1140
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":11F3
         Caption         =   "frmYpmf110.frx":1213
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1281
         Keys            =   "frmYpmf110.frx":129F
         Spin            =   "frmYpmf110.frx":12E9
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   1620
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "雑収入"
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
         LabelLeft       =   36
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
      Begin imNumber6Ctl.imNumber imnOtotal 
         Height          =   435
         Left            =   1920
         TabIndex        =   4
         Top             =   1620
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":1311
         Caption         =   "frmYpmf110.frx":1331
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":139F
         Keys            =   "frmYpmf110.frx":13BD
         Spin            =   "frmYpmf110.frx":1407
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   44
         Top             =   2640
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "維持管理費"
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
         LabelLeft       =   21
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
      Begin imNumber6Ctl.imNumber imnRkeep 
         Height          =   435
         Left            =   1920
         TabIndex        =   6
         Top             =   3060
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":142F
         Caption         =   "frmYpmf110.frx":144F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":14BD
         Keys            =   "frmYpmf110.frx":14DB
         Spin            =   "frmYpmf110.frx":1525
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnBkeep 
         Height          =   435
         Left            =   1920
         TabIndex        =   7
         Top             =   3540
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":154D
         Caption         =   "frmYpmf110.frx":156D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":15DB
         Keys            =   "frmYpmf110.frx":15F9
         Spin            =   "frmYpmf110.frx":1643
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnSkeep 
         Height          =   435
         Left            =   1920
         TabIndex        =   8
         Top             =   4020
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":166B
         Caption         =   "frmYpmf110.frx":168B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":16F9
         Keys            =   "frmYpmf110.frx":1717
         Spin            =   "frmYpmf110.frx":1761
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   45
         Top             =   4560
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "買主消費税"
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
         LabelLeft       =   21
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
      Begin imNumber6Ctl.imNumber imnBtax 
         Height          =   435
         Left            =   1920
         TabIndex        =   9
         Top             =   4560
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":1789
         Caption         =   "frmYpmf110.frx":17A9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1817
         Keys            =   "frmYpmf110.frx":1835
         Spin            =   "frmYpmf110.frx":187F
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   46
         Top             =   5040
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "仕入売上分"
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
         LabelLeft       =   21
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
      Begin imNumber6Ctl.imNumber imnPsall 
         Height          =   435
         Left            =   1920
         TabIndex        =   14
         Top             =   5040
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":18A7
         Caption         =   "frmYpmf110.frx":18C7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1935
         Keys            =   "frmYpmf110.frx":1953
         Spin            =   "frmYpmf110.frx":199D
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
         ValueVT         =   5
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   47
         Top             =   8160
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "小計①"
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
         LabelLeft       =   36
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
      Begin imNumber6Ctl.imNumber imnTotal1 
         Height          =   435
         Left            =   1920
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   8160
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":19C5
         Caption         =   "frmYpmf110.frx":19E5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1A53
         Keys            =   "frmYpmf110.frx":1A71
         Spin            =   "frmYpmf110.frx":1ABB
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1179653
         Value           =   9999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   5
         Left            =   6240
         TabIndex        =   48
         Top             =   180
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "経　費"
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
         LabelWidth      =   40
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
      Begin imNumber6Ctl.imNumber imnExpenses 
         Height          =   435
         Left            =   8040
         TabIndex        =   17
         Top             =   180
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":1AE3
         Caption         =   "frmYpmf110.frx":1B03
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1B71
         Keys            =   "frmYpmf110.frx":1B8F
         Spin            =   "frmYpmf110.frx":1BD9
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   6
         Left            =   6240
         TabIndex        =   49
         Top             =   1140
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "出品者消費税"
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
         LabelWidth      =   90
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
      Begin imNumber6Ctl.imNumber imnStax 
         Height          =   435
         Left            =   8040
         TabIndex        =   19
         Top             =   1140
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":1C01
         Caption         =   "frmYpmf110.frx":1C21
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1C8F
         Keys            =   "frmYpmf110.frx":1CAD
         Spin            =   "frmYpmf110.frx":1CF7
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   9
         Left            =   6240
         TabIndex        =   50
         Top             =   1620
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "仕　入"
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
         LabelWidth      =   40
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
      Begin imNumber6Ctl.imNumber imnPtotal 
         Height          =   435
         Left            =   8040
         TabIndex        =   20
         Top             =   1620
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":1D1F
         Caption         =   "frmYpmf110.frx":1D3F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1DAD
         Keys            =   "frmYpmf110.frx":1DCB
         Spin            =   "frmYpmf110.frx":1E15
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   13
         Left            =   6240
         TabIndex        =   51
         Top             =   2100
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "今回未収金"
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
         LabelLeft       =   21
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
      Begin imNumber6Ctl.imNumber imnUtotal 
         Height          =   435
         Left            =   8040
         TabIndex        =   21
         Top             =   2100
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":1E3D
         Caption         =   "frmYpmf110.frx":1E5D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1ECB
         Keys            =   "frmYpmf110.frx":1EE9
         Spin            =   "frmYpmf110.frx":1F33
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   14
         Left            =   6240
         TabIndex        =   52
         Top             =   3240
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "小計②"
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
         LabelLeft       =   36
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
      Begin imNumber6Ctl.imNumber imnTotal2 
         Height          =   435
         Left            =   8040
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":1F5B
         Caption         =   "frmYpmf110.frx":1F7B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":1FE9
         Keys            =   "frmYpmf110.frx":2007
         Spin            =   "frmYpmf110.frx":2051
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1179653
         Value           =   9999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   15
         Left            =   6240
         TabIndex        =   53
         Top             =   4020
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "小計①－小計②"
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
         LabelWidth      =   105
         LabelHeight     =   25
         LabelLeft       =   6
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
      Begin imNumber6Ctl.imNumber imnGtotal 
         Height          =   555
         Left            =   8040
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4020
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   979
         Calculator      =   "frmYpmf110.frx":2079
         Caption         =   "frmYpmf110.frx":2099
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":2107
         Keys            =   "frmYpmf110.frx":2125
         Spin            =   "frmYpmf110.frx":216F
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1179653
         Value           =   9999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   58
         Top             =   2160
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "造園工事売上金"
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
         LabelWidth      =   105
         LabelHeight     =   25
         LabelLeft       =   6
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
      Begin imNumber6Ctl.imNumber imnGtotalSales 
         Height          =   435
         Left            =   1920
         TabIndex        =   5
         Top             =   2160
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":2197
         Caption         =   "frmYpmf110.frx":21B7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":2225
         Keys            =   "frmYpmf110.frx":2243
         Spin            =   "frmYpmf110.frx":228D
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   17
         Left            =   6240
         TabIndex        =   61
         Top             =   660
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "造園工事仕入金"
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
         LabelWidth      =   105
         LabelHeight     =   25
         LabelLeft       =   6
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
      Begin imNumber6Ctl.imNumber imnGtotalStock 
         Height          =   435
         Left            =   8040
         TabIndex        =   18
         Top             =   660
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":22B5
         Caption         =   "frmYpmf110.frx":22D5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":2343
         Keys            =   "frmYpmf110.frx":2361
         Spin            =   "frmYpmf110.frx":23AB
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   18
         Left            =   120
         TabIndex        =   72
         Top             =   5520
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "自社注文出品分"
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
         LabelWidth      =   105
         LabelHeight     =   25
         LabelLeft       =   6
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
      Begin imNumber6Ctl.imNumber imnJctotal 
         Height          =   435
         Left            =   1920
         TabIndex        =   16
         Top             =   5520
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":23D3
         Caption         =   "frmYpmf110.frx":23F3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":2461
         Keys            =   "frmYpmf110.frx":247F
         Spin            =   "frmYpmf110.frx":24C9
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
         ValueVT         =   5
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   19
         Left            =   6240
         TabIndex        =   74
         Top             =   2580
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "自社競売購入分"
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
         LabelWidth      =   105
         LabelHeight     =   25
         LabelLeft       =   6
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
      Begin imNumber6Ctl.imNumber imnJbtotal 
         Height          =   435
         Left            =   8040
         TabIndex        =   22
         Top             =   2580
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":24F1
         Caption         =   "frmYpmf110.frx":2511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":257F
         Keys            =   "frmYpmf110.frx":259D
         Spin            =   "frmYpmf110.frx":25E7
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   20
         Left            =   120
         TabIndex        =   76
         Top             =   6000
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "買主競売手数料"
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
         LabelWidth      =   105
         LabelHeight     =   25
         LabelLeft       =   6
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
      Begin imNumber6Ctl.imNumber imnBrate2 
         Height          =   435
         Left            =   1920
         TabIndex        =   10
         Top             =   6000
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":260F
         Caption         =   "frmYpmf110.frx":262F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":269D
         Keys            =   "frmYpmf110.frx":26BB
         Spin            =   "frmYpmf110.frx":2705
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   21
         Left            =   120
         TabIndex        =   78
         Top             =   6480
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "受付伝票代"
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
         LabelLeft       =   21
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
      Begin imNumber6Ctl.imNumber imnRrate 
         Height          =   435
         Left            =   1920
         TabIndex        =   11
         Top             =   6480
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":272D
         Caption         =   "frmYpmf110.frx":274D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":27BB
         Keys            =   "frmYpmf110.frx":27D9
         Spin            =   "frmYpmf110.frx":2823
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
         ValueVT         =   5
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   22
         Left            =   120
         TabIndex        =   79
         Top             =   6960
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "荷札代"
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
         LabelLeft       =   36
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
         Index           =   23
         Left            =   120
         TabIndex        =   80
         Top             =   7440
         Width           =   1755
         _Version        =   262145
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "印紙代"
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
         LabelLeft       =   36
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
      Begin imNumber6Ctl.imNumber imnEf 
         Height          =   435
         Left            =   1920
         TabIndex        =   12
         Top             =   6960
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":284B
         Caption         =   "frmYpmf110.frx":286B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":28D9
         Keys            =   "frmYpmf110.frx":28F7
         Spin            =   "frmYpmf110.frx":2941
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnStamp 
         Height          =   435
         Left            =   1920
         TabIndex        =   13
         Top             =   7440
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   767
         Calculator      =   "frmYpmf110.frx":2969
         Caption         =   "frmYpmf110.frx":2989
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf110.frx":29F7
         Keys            =   "frmYpmf110.frx":2A15
         Spin            =   "frmYpmf110.frx":2A5F
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
         ValueVT         =   1179653
         Value           =   999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "※未収分も含む"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   18
         Left            =   4140
         TabIndex        =   77
         Top             =   4680
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "←入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Index           =   17
         Left            =   10200
         TabIndex        =   75
         Top             =   2700
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "←入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Index           =   15
         Left            =   4140
         TabIndex        =   73
         Top             =   5640
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "←入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Index           =   16
         Left            =   10260
         TabIndex        =   71
         Top             =   1680
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "←入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Index           =   14
         Left            =   10260
         TabIndex        =   70
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "←入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Index           =   13
         Left            =   10260
         TabIndex        =   69
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "←入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Index           =   12
         Left            =   4140
         TabIndex        =   68
         Top             =   2220
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "←入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Index           =   11
         Left            =   4140
         TabIndex        =   67
         Top             =   1740
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "←入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Index           =   10
         Left            =   4140
         TabIndex        =   66
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "※注文分"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   9
         Left            =   4140
         TabIndex        =   65
         Top             =   5160
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "※入金分"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   8
         Left            =   4140
         TabIndex        =   64
         Top             =   3180
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "※維持管理費,手数料を除く"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   7
         Left            =   4140
         TabIndex        =   63
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "※維持管理費と消費税を含む"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   6
         Left            =   10260
         TabIndex        =   62
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "※未収分も含む"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   5
         Left            =   4140
         TabIndex        =   60
         Top             =   6120
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "※未収分も含む"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   4
         Left            =   4140
         TabIndex        =   59
         Top             =   3600
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "※出品者手数料"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   3
         Left            =   4140
         TabIndex        =   57
         Top             =   780
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "出品"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   2
         Left            =   1140
         TabIndex        =   56
         Top             =   4080
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "買主"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   1140
         TabIndex        =   55
         Top             =   3600
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "未収"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   1140
         TabIndex        =   54
         Top             =   3120
         Width           =   645
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   6240
         X2              =   10320
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   6240
         X2              =   10200
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   60
         X2              =   4080
         Y1              =   8040
         Y2              =   8040
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   12600
      TabIndex        =   0
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf110.frx":2A87
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf110.frx":2AF5
      Key             =   "frmYpmf110.frx":2B13
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
      Left            =   12600
      TabIndex        =   29
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf110.frx":2B57
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf110.frx":2BC5
      Key             =   "frmYpmf110.frx":2BE3
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
      TabIndex        =   30
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf110.frx":2C27
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf110.frx":2C95
      Key             =   "frmYpmf110.frx":2CB3
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
Attribute VB_Name = "frmYpmf110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()

    If MsgBox("再計算しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    If FieldsCalc = False Then Exit Sub
    Call Calc_Total

End Sub

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    imnCashe.SetFocus

End Sub

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If MsgBox("実行しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    '入力チェック
    If DoValidationChecks() = False Then Exit Sub
    'データ更新
    If DataUpdate = False Then Exit Sub
    '印刷用ワーク作成
    If MakePrintWork = False Then Exit Sub
    '印刷プレビュー
    If ActiveReportPrint(0) = False Then Exit Sub

    Exit Sub
    
cmdExecute_Click_Err:

    Call MsgBox("実行クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

'目　的　　：
'条　件　　：終了クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
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

    Call MsgBox("開催年月日と担当者の変更クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdLogin_Click_Err")

End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
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
            cmdCalc.SetFocus
            DoEvents
            Call cmdCalc_Click
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

    Call MsgBox("フォームキーダウン時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")

End Sub

'目　的　　：
'条　件　　：フォームロード時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "日別日計表"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
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
'作成年月日：２００２／０７／２５
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
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        imnCashe.Value = 0
        imnProfit.Value = 0
        imnRtotal.Value = 0
        imnOtotal.Value = 0
        imnGtotalSales.Value = 0
        imnRkeep.Value = 0
        imnBkeep.Value = 0
        imnSkeep.Value = 0
        imnBtax.Value = 0
        imnPsall.Value = 0
        imnTotal1.Value = 0
        imnExpenses.Value = 0
        imnGtotalStock.Value = 0
        imnStax.Value = 0
        imnPtotal.Value = 0
        imnUtotal.Value = 0
        imnTotal2.Value = 0
        imnGtotal.Value = 0
        imnJctotal.Value = 0
        imnJbtotal.Value = 0
        
        '201107
        imnBrate2.Value = 0
        imnRrate.Value = 0
        imnEf.Value = 0
        imnStamp.Value = 0
        
        Call FieldsSet
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
'作成年月日：２００２／０７／２５
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

'目　的　　：印刷用ワーク作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim wkRecordset As New ADODB.Recordset

    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF110"
    g_clsAdoAccess.Connection.Execute strSQL

    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF110"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    wkRecordset.AddNew
    wkRecordset.Fields("Odate") = lblOdate.Caption
    wkRecordset.Fields("Cashe") = imnCashe.Value
    wkRecordset.Fields("Profit") = imnProfit.Value
    wkRecordset.Fields("Rtotal") = imnRtotal.Value
    wkRecordset.Fields("Ototal") = imnOtotal.Value
    wkRecordset.Fields("GtotalSales") = imnGtotalSales.Value
    wkRecordset.Fields("Rkeep") = imnRkeep.Value
    wkRecordset.Fields("Bkeep") = imnBkeep.Value
    wkRecordset.Fields("Skeep") = imnSkeep.Value
    wkRecordset.Fields("Btax") = imnBtax.Value
    wkRecordset.Fields("Psall") = imnPsall.Value
    wkRecordset.Fields("Total1") = imnTotal1.Value
    wkRecordset.Fields("Expenses") = imnExpenses.Value
    wkRecordset.Fields("GtotalStock") = imnGtotalStock.Value
    wkRecordset.Fields("Stax") = imnStax.Value
    wkRecordset.Fields("Ptotal") = imnPtotal.Value
    wkRecordset.Fields("Utotal") = imnUtotal.Value
    wkRecordset.Fields("Total2") = imnTotal2.Value
    wkRecordset.Fields("Gtotal") = imnGtotal.Value
    wkRecordset.Fields("Jctotal") = imnJctotal.Value
    wkRecordset.Fields("Jbtotal") = imnJbtotal.Value
    
    '201107
    wkRecordset.Fields("Brate2") = imnBrate2.Value
    wkRecordset.Fields("Rrate") = imnRrate.Value
    wkRecordset.Fields("Ef") = imnEf.Value
    wkRecordset.Fields("Stamp") = imnStamp.Value
    
    wkRecordset.Update
    
    wkRecordset.Requery
    wkRecordset.Close
    
    MakePrintWork = True
    
MakePrintWork_Exit:
    
    Screen.MousePointer = vbDefault
    
    Exit Function

MakePrintWork_Cancel:

    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    MakePrintWork = False
    Call MsgBox("印刷ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function

Private Sub imnBkeep_GotFocus()
    
    imnBkeep.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnBkeep_LostFocus()
    
    Call Calc_Total
    imnBkeep.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnBtax_GotFocus()
    
    imnBtax.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnBtax_LostFocus()
    
    Call Calc_Total
    imnBtax.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnCashe_GotFocus()
    
    imnCashe.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnCashe_LostFocus()
    
    Call Calc_Total
    imnCashe.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnExpenses_GotFocus()
    
    imnExpenses.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnExpenses_LostFocus()
    
    Call Calc_Total
    imnExpenses.BackColor = FOCUS_NO_COLOR
 
End Sub

Private Sub imnGtotalSales_GotFocus()
    
    imnGtotalSales.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnGtotalSales_LostFocus()
    
    Call Calc_Total
    imnGtotalSales.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnGtotalStock_GotFocus()
    
    imnGtotalStock.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnGtotalStock_LostFocus()
    
    Call Calc_Total
    imnGtotalStock.BackColor = FOCUS_NO_COLOR
 
End Sub

Private Sub imnJbtotal_GotFocus()
    
    imnJbtotal.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnJbtotal_LostFocus()
    
    Call Calc_Total
    imnJbtotal.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnJctotal_GotFocus()
    
    imnJctotal.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnJctotal_LostFocus()
    
    Call Calc_Total
    imnJctotal.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnOtotal_GotFocus()
    
    imnOtotal.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnOtotal_LostFocus()
    
    Call Calc_Total
    imnOtotal.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnProfit_GotFocus()
    
    imnProfit.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnProfit_LostFocus()
    
    Call Calc_Total
    imnProfit.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnPsall_GotFocus()
    
    imnPsall.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnPsall_LostFocus()
    
    Call Calc_Total
    imnPsall.BackColor = FOCUS_NO_COLOR
 
End Sub

Private Sub imnPtotal_GotFocus()
    
    imnPtotal.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnPtotal_LostFocus()
    
    Call Calc_Total
    imnPtotal.BackColor = FOCUS_NO_COLOR
 
End Sub

Private Sub imnRkeep_GotFocus()
    
    imnRkeep.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRkeep_LostFocus()
    
    Call Calc_Total
    imnRkeep.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnBrate2_GotFocus()
    
    imnBrate2.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBrate2_LostFocus()
    
    Call Calc_Total
    imnBrate2.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRtotal_GotFocus()
    
    imnRtotal.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRtotal_LostFocus()
    
    Call Calc_Total
    imnRtotal.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnSkeep_GotFocus()
    
    imnSkeep.BackColor = FOCUS_STOP_COLOR
  
End Sub

Private Sub imnSkeep_LostFocus()
    
    Call Calc_Total
    imnSkeep.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnStax_GotFocus()
    
    imnStax.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnStax_LostFocus()
    
    Call Calc_Total
    imnStax.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnTotal1_LostFocus()
    
'    Call Calc_Total
    
End Sub

Private Sub imnTotal2_LostFocus()
    
'    Call Calc_Total
    
End Sub

Private Sub imnUtotal_GotFocus()
    
    imnUtotal.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnUtotal_LostFocus()
    
    Call Calc_Total
    imnUtotal.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    imnCashe.SetFocus

End Sub

Private Sub imnRrate_GotFocus()
    
    imnRrate.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnRrate_LostFocus()
    
    Call Calc_Total
    imnRrate.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnEf_GotFocus()
    
    imnEf.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnEf_LostFocus()
    
    Call Calc_Total
    imnEf.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnStamp_GotFocus()
    
    imnStamp.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub imnStamp_LostFocus()
    
    Call Calc_Total
    imnStamp.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：ActiveReportの印刷
'条　件　　：
'結　果　　：
'引　数　　：0:プレビュー 1:印刷
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf110
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "日別日計表"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .Caption = "日別日計表"
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
    Call MsgBox("実行クリックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReportPrint_Err")
    
End Function

'目　的　　：フィールドのセット
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Function FieldsSet() As Boolean
    
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    '開催日別日計
    strSQL = "SELECT * FROM DT050" & _
             " WHERE Odate = '" & lblOdate.Caption & "'"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        imnCashe.Value = IIf(IsNull(adoRecordset1.Fields("Cashe")), 0, adoRecordset1.Fields("Cashe"))
        imnProfit.Value = IIf(IsNull(adoRecordset1.Fields("Profit")), 0, adoRecordset1.Fields("Profit"))
        imnRtotal.Value = IIf(IsNull(adoRecordset1.Fields("Rtotal")), 0, adoRecordset1.Fields("Rtotal"))
        imnOtotal.Value = IIf(IsNull(adoRecordset1.Fields("Ototal")), 0, adoRecordset1.Fields("Ototal"))
        imnRkeep.Value = IIf(IsNull(adoRecordset1.Fields("Rkeep")), 0, adoRecordset1.Fields("Rkeep"))
        imnBkeep.Value = IIf(IsNull(adoRecordset1.Fields("Bkeep")), 0, adoRecordset1.Fields("Bkeep"))
        imnSkeep.Value = IIf(IsNull(adoRecordset1.Fields("Skeep")), 0, adoRecordset1.Fields("Skeep"))
        imnBtax.Value = IIf(IsNull(adoRecordset1.Fields("Btax")), 0, adoRecordset1.Fields("Btax"))
        imnPsall.Value = IIf(IsNull(adoRecordset1.Fields("Psall")), 0, adoRecordset1.Fields("Psall"))
        imnExpenses.Value = IIf(IsNull(adoRecordset1.Fields("Expenses")), 0, adoRecordset1.Fields("Expenses"))
        imnStax.Value = IIf(IsNull(adoRecordset1.Fields("Stax")), 0, adoRecordset1.Fields("Stax"))
        imnPtotal.Value = IIf(IsNull(adoRecordset1.Fields("Ptotal")), 0, adoRecordset1.Fields("Ptotal"))
        imnUtotal.Value = IIf(IsNull(adoRecordset1.Fields("Utotal")), 0, adoRecordset1.Fields("Utotal"))
        imnJctotal.Value = IIf(IsNull(adoRecordset1.Fields("Jctotal")), 0, adoRecordset1.Fields("Jctotal"))
        imnJbtotal.Value = IIf(IsNull(adoRecordset1.Fields("Jbtotal")), 0, adoRecordset1.Fields("Jbtotal"))
        
        '201107
        imnBrate2.Value = IIf(IsNull(adoRecordset1.Fields("Brate2")), 0, adoRecordset1.Fields("Brate2"))
        imnRrate.Value = IIf(IsNull(adoRecordset1.Fields("Rrate")), 0, adoRecordset1.Fields("Rrate"))
        imnEf.Value = IIf(IsNull(adoRecordset1.Fields("Ef")), 0, adoRecordset1.Fields("Ef"))
        imnStamp.Value = IIf(IsNull(adoRecordset1.Fields("Stamp")), 0, adoRecordset1.Fields("Stamp"))
        
        
    Else
        '集計フィールドのセット
        Call FieldsCalc
    End If
    
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
    '小計などの計算
    Call Calc_Total
    
    Screen.MousePointer = vbDefault
    
    FieldsSet = True
    
    Exit Function

FieldsSet_Err:

    FieldsSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("フィールドセットエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsSet_Err")

End Function

'目　的　　：データの登録
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
    With adoRecordset1
        strSQL = "SELECT * FROM DT050" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        .Fields("Odate") = lblOdate.Caption
        .Fields("Cashe") = imnCashe.Value
        .Fields("Profit") = imnProfit.Value
        .Fields("Rtotal") = imnRtotal.Value
        .Fields("Ototal") = imnOtotal.Value
        .Fields("GtotalSales") = imnGtotalSales.Value
        .Fields("Rkeep") = imnRkeep.Value
        .Fields("Bkeep") = imnBkeep.Value
        .Fields("Skeep") = imnSkeep.Value
        .Fields("Btax") = imnBtax.Value
        .Fields("Psall") = imnPsall.Value
        .Fields("Expenses") = imnExpenses.Value
        .Fields("GtotalStock") = imnGtotalStock.Value
        .Fields("Stax") = imnStax.Value
        .Fields("Ptotal") = imnPtotal.Value
        .Fields("Utotal") = imnUtotal.Value
        .Fields("Jctotal") = imnJctotal.Value
        .Fields("Jbtotal") = imnJbtotal.Value
        
        '201107
        .Fields("Brate2") = imnBrate2.Value
        .Fields("Rrate") = imnRrate.Value
        .Fields("Ef") = imnEf.Value
        .Fields("Stamp") = imnStamp.Value
        
        .Update
        .Close
        
    End With
    
    g_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    Set adoRecordset1 = Nothing
    
    DataUpdate = True
    
    Exit Function

DataUpdate_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    DataUpdate = False
    Screen.MousePointer = vbDefault
    Call MsgBox("データ登録エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_Err")

End Function

'目　的　　：小計などの計算
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Sub Calc_Total()

    On Error GoTo Calc_Total_Err
    
    '小計①の計算
    imnTotal1.Value = CCur(imnCashe.Value) + CCur(imnProfit.Value) _
                    + CCur(imnRtotal.Value) + CCur(imnOtotal.Value) + CCur(imnGtotalSales.Value) _
                    + CCur(imnRkeep.Value) + CCur(imnBkeep.Value) _
                    + CCur(imnSkeep.Value) + CCur(imnBtax.Value) _
                    + CCur(imnPsall.Value) + CCur(imnJctotal.Value) _
                    + CCur(imnBrate2.Value) + CCur(imnRrate.Value) + CCur(imnEf.Value) + CCur(imnStamp.Value) '201107'201107
    '小計②の計算
    imnTotal2.Value = CCur(imnExpenses.Value) + CCur(imnGtotalStock.Value) + CCur(imnStax.Value) _
                    + CCur(imnPtotal.Value) + CCur(imnUtotal.Value) + CCur(imnJbtotal.Value)
    
    '小計①－小計②
    imnGtotal.Value = CCur(imnTotal1.Value) - CCur(imnTotal2.Value)
    
    Exit Sub
    
Calc_Total_Err:

    Call MsgBox("合計の計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Total_Err")

End Sub

'目　的　　：集計フィールドのセット
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２５
'更新履歴　：
'
Private Function FieldsCalc() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim adoMT010 As New ADODB.Recordset
    Dim curTotal1 As Currency
    Dim curTotal2 As Currency
    
    On Error GoTo FieldsCalc_Err

    FieldsCalc = False

'********** 利益 **********

    '出品者精算データ
    strSQL = "{call sp_YPMF1101;1('" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        imnProfit.Value = adoRecordset1.Fields("Charge_Total")
    Else
        imnProfit.Value = 0
    End If
    adoRecordset1.Close

'********** 未収分の入金 **********

    '入金日が開催日と等しく、現金入金データが対象
    '入金データ
    strSQL = "{call sp_YPMF1108;1('" & lblOdate.Caption & "','" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        If Not IsNull(adoRecordset1.Fields("Ptotal")) And Not IsNull(adoRecordset1.Fields("Ptotal2")) Then
            '201107
            imnRtotal.Value = CCur(adoRecordset1.Fields("Ptotal")) - CCur(adoRecordset1.Fields("Ptotal2")) - CCur(adoRecordset1.Fields("Ptotal3"))
        Else
            imnRtotal.Value = 0
        End If
    Else
        imnRtotal.Value = 0
    End If
    adoRecordset1.Close

'********** 維持管理費(未収) **********

    '入金日が開催日と等しく、現金入金データが対象
    '入金データ
    strSQL = "{call sp_YPMF1107;1('" & lblOdate.Caption & "','" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        If Not IsNull(adoRecordset1.Fields("Ptotal2")) Then
            imnRkeep.Value = adoRecordset1.Fields("Ptotal2")
        Else
            imnRkeep.Value = 0
        End If
    Else
        imnRkeep.Value = 0
    End If
    adoRecordset1.Close

        '201107
'********** 維持管理費,競売手数料(買主) **********

    '買主精算データ
    strSQL = "{call sp_YPMF1102;1('" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        imnBkeep.Value = adoRecordset1.Fields("Keep_Total")
        '201107
        imnBrate2.Value = adoRecordset1.Fields("Brate2_Total")
    Else
        imnBkeep.Value = 0
        '201107
        imnBrate2.Value = 0
    End If
    adoRecordset1.Close


    '入金日が開催日と等しく、現金入金データが対象 買主競売手数料を足す 2012/06/16
    strSQL = "SELECT SUM(Ptotal3) AS Ptotal3 From DT060 WHERE R = 1 AND (Rdate='" & lblOdate.Caption & "') "
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        '201107
        If Not IsNull(adoRecordset1.Fields("Ptotal3")) Then
            imnBrate2.Value = imnBrate2.Value + adoRecordset1.Fields("Ptotal3")
        End If
    End If
    adoRecordset1.Close


        '201107
'********** 維持管理費,受付伝票代,荷札代,印紙代(出品者) **********

    '出品者精算データ
    strSQL = "{call sp_YPMF1103;1('" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        imnSkeep.Value = adoRecordset1.Fields("Keep_Total")
        '201107
        imnRrate.Value = adoRecordset1.Fields("Rrate_Total")
        imnEf.Value = adoRecordset1.Fields("Ef_Total")
        imnStamp.Value = adoRecordset1.Fields("Stamp_Total")
    Else
        imnSkeep.Value = 0
        '201107
        imnRrate.Value = 0
        imnEf.Value = 0
        imnStamp.Value = 0
    End If
    adoRecordset1.Close

'********** 買主消費税 **********

    '買主精算データ
    strSQL = "{call sp_YPMF1104;1('" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        imnBtax.Value = adoRecordset1.Fields("Tax_Total")
    Else
        imnBtax.Value = 0
    End If
    adoRecordset1.Close

'********** 出品者消費税 **********

    '出品者精算データ
    strSQL = "{call sp_YPMF1105;1('" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        imnStax.Value = adoRecordset1.Fields("Tax_Total")
    Else
        imnStax.Value = 0
    End If
    adoRecordset1.Close

'********** 今回未収金 **********

    '出品者精算データ
    strSQL = "{call sp_YPMF1106;1('" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        imnUtotal.Value = adoRecordset1.Fields("Gtotal_Total")
    Else
        imnUtotal.Value = 0
    End If
    adoRecordset1.Close

'********** 仕入売上分 **********

    Dim curPrice As Currency
    Dim curCharge As Currency
    Dim curTotal As Currency
    Dim curTax As Currency
    Dim curKeep As Currency
    Dim curGtotal As Currency
    Dim curSkeep As Currency                '出品者維持管理費
    Dim curSrate As Currency                '出品者手数料率
    Dim intSfraction As Integer             '出品者端数処理
    Dim curSRounding As Currency            '出品者丸め単位
    Dim curTaxRate As Currency              '消費税率

    '初期化
    curSkeep = 0        '出品者維持管理費
    intSfraction = 0    '出品者端数処理
    curSrate = 0        '出品者手数料率
    curSRounding = 0    '出品者丸め単位
    curTaxRate = 0      '消費税率
    
    '設定マスタオープン
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        If Not IsNull(adoMT010.Fields("Skeep")) Then curSkeep = adoMT010.Fields("Skeep")
        If Not IsNull(adoMT010.Fields("Sfraction")) Then intSfraction = adoMT010.Fields("Sfraction")
        If Not IsNull(adoMT010.Fields("Srate")) Then curSrate = adoMT010.Fields("Srate")
        If Not IsNull(adoMT010.Fields("SRounding")) Then curSRounding = adoMT010.Fields("SRounding")
    End If
    adoMT010.Close
    
    '消費税率取得
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, g_strOdate)

    curTotal1 = 0
    curTotal2 = 0

    '買主精算データの売立合計
    strSQL = "{call sp_YPMF1109;1('" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        If IsNull(adoRecordset1.Fields("Total")) = False Then
            curTotal1 = adoRecordset1.Fields("Total")
        End If
    End If
    adoRecordset1.Close

    '注文データの総合計取得
    strSQL = "{call sp_YPMF1109;2('" & lblOdate.Caption & "')}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        '初期化
        curPrice = 0
        curCharge = 0
        curTax = 0
        curKeep = 0
        curGtotal = 0
        
        '注文明細データ
        strSQL = "{call sp_YPMF1109;3('" & lblOdate.Caption & "'," & adoRecordset1.Fields("Onum") & ")}"
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset2.EOF
            '仕入金額＝数量×仕入単価
            curPrice = curPrice + (CCur(adoRecordset2.Fields("Qty")) * CCur(adoRecordset2.Fields("Price2")))

            adoRecordset2.MoveNext
        Loop
        adoRecordset2.Close
        
        '手数料
        If IsNull(adoRecordset1.Fields("ChargeDiv")) = False And adoRecordset1.Fields("ChargeDiv") = 1 Then
            curCharge = curPrice * curSrate / 100
            '切り捨て
            If IsNull(adoRecordset1.Fields("FixDiv")) = False And adoRecordset1.Fields("FixDiv") = 1 Then
                curCharge = Global_Rounding(curCharge, intSfraction, curSRounding)
            Else
                curCharge = Global_Rounding(curCharge, intSfraction, 1)
            End If
        End If
        '消費税計算
        If IsNull(adoRecordset1.Fields("TaxDiv")) = False And adoRecordset1.Fields("TaxDiv") = 1 Then
            '切り捨て
            If IsNull(adoRecordset1.Fields("FixDiv")) = False And adoRecordset1.Fields("FixDiv") = 1 Then
                curTax = Global_Get_Tax(curPrice - curCharge, curTaxRate, intSfraction, curSRounding)
            Else
                curTax = Global_Get_Tax(curPrice - curCharge, curTaxRate, intSfraction, 1)
            End If
        End If
        '維持管理費
        If IsNull(adoRecordset1.Fields("KeepDiv")) = False And adoRecordset1.Fields("KeepDiv") = 1 Then
            curKeep = curSkeep
        End If
        '総合計＝合計－手数料＋消費税－維持管理費
        curGtotal = curPrice - curCharge + curTax - curKeep

        curTotal2 = curTotal2 + curGtotal
        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close

    imnPsall.Value = curTotal1 - curTotal2

    FieldsCalc = True

    Exit Function

FieldsCalc_Err:

    FieldsCalc = False
    Call MsgBox("集計フィールドのセットエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsCalc_Err")

End Function
