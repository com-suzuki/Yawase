VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmDt041 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDt041.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10740
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraRecordSelector 
      Height          =   615
      Left            =   7740
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton cmdDataMove 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Picture         =   "frmDt041.frx":0CFA
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Picture         =   "frmDt041.frx":0E44
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Picture         =   "frmDt041.frx":0F8E
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Picture         =   "frmDt041.frx":10D8
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
   End
   Begin VB.Frame fra 
      Height          =   735
      Left            =   60
      TabIndex        =   29
      Top             =   5760
      Width           =   10575
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         Height          =   495
         Left            =   60
         TabIndex        =   16
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
         Picture         =   "frmDt041.frx":1222
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8760
         TabIndex        =   18
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
         Picture         =   "frmDt041.frx":123E
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   7020
         TabIndex        =   17
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "実行(F12)"
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
         Picture         =   "frmDt041.frx":1398
      End
   End
   Begin VB.Frame fraSyori 
      Height          =   615
      Left            =   60
      TabIndex        =   23
      Top             =   0
      Width           =   7635
      Begin VB.OptionButton optSyori 
         Caption         =   "外部出力"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "印　刷"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "削　除"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "変　更"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "新　規"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   27
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "処理区分"
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
   End
   Begin VB.Frame fraMeisai 
      Height          =   4035
      Left            =   60
      TabIndex        =   21
      Top             =   1740
      Width           =   10575
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'なし
         Height          =   615
         Left            =   4680
         TabIndex        =   50
         Top             =   1680
         Width           =   4335
         Begin VB.OptionButton optR 
            Caption         =   "小切手"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1440
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   11
            Top             =   180
            Width           =   1395
         End
         Begin VB.OptionButton optR 
            Caption         =   "現　金"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   10
            Top             =   180
            Width           =   1395
         End
         Begin VB.OptionButton optR 
            Caption         =   "銀行振込"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2880
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   12
            Top             =   180
            Width           =   1395
         End
      End
      Begin VB.CheckBox chkRdiv 
         Caption         =   "入金済"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   9
         Top             =   1860
         Width           =   1215
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   3
         Left            =   60
         TabIndex        =   28
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "合計金額"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   38
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "消費税額"
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
      Begin imNumber6Ctl.imNumber imnTotal 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   180
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":17EA
         Caption         =   "frmDt041.frx":180A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":1878
         Keys            =   "frmDt041.frx":1896
         Spin            =   "frmDt041.frx":18E0
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
      Begin imNumber6Ctl.imNumber imnTax 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":1908
         Caption         =   "frmDt041.frx":1928
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":1996
         Keys            =   "frmDt041.frx":19B4
         Spin            =   "frmDt041.frx":19FE
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
      Begin imNumber6Ctl.imNumber imnKeep 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1020
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":1A26
         Caption         =   "frmDt041.frx":1A46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":1AB4
         Keys            =   "frmDt041.frx":1AD2
         Spin            =   "frmDt041.frx":1B1C
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
      Begin imNumber6Ctl.imNumber imnGtotal 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":1B44
         Caption         =   "frmDt041.frx":1B64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":1BD2
         Keys            =   "frmDt041.frx":1BF0
         Spin            =   "frmDt041.frx":1C3A
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
         Left            =   60
         TabIndex        =   40
         Top             =   1020
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
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
         Index           =   6
         Left            =   60
         TabIndex        =   41
         Top             =   1440
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "総合計金額"
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
         Index           =   7
         Left            =   60
         TabIndex        =   42
         Top             =   1860
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "入金区分"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   8
         Left            =   3180
         TabIndex        =   43
         Top             =   1860
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   9
         Left            =   60
         TabIndex        =   44
         Top             =   2340
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "入金年月日"
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
         Index           =   10
         Left            =   60
         TabIndex        =   45
         Top             =   3540
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "入力時間"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   11
         Left            =   60
         TabIndex        =   46
         Top             =   3120
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "担当者コード"
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
         LabelWidth      =   82
         LabelHeight     =   25
         LabelLeft       =   7
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
         Left            =   1560
         TabIndex        =   13
         Top             =   2340
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":1C62
         Caption         =   "frmDt041.frx":1C82
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":1CF0
         Keys            =   "frmDt041.frx":1D0E
         Spin            =   "frmDt041.frx":1D48
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
         ValueVT         =   2011496453
         Value           =   2099
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnRdate_Month 
         Height          =   375
         Left            =   2700
         TabIndex        =   14
         Top             =   2340
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":1D70
         Caption         =   "frmDt041.frx":1D90
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":1DFE
         Keys            =   "frmDt041.frx":1E1C
         Spin            =   "frmDt041.frx":1E56
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
         ValueVT         =   2011496453
         Value           =   12
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnRdate_Day 
         Height          =   375
         Left            =   3540
         TabIndex        =   15
         Top             =   2340
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":1E7E
         Caption         =   "frmDt041.frx":1E9E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":1F0C
         Keys            =   "frmDt041.frx":1F2A
         Spin            =   "frmDt041.frx":1F64
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
         ValueVT         =   2011496453
         Value           =   31
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnBrate2 
         Height          =   375
         Left            =   4620
         TabIndex        =   7
         Top             =   1020
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":1F8C
         Caption         =   "frmDt041.frx":1FAC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":201A
         Keys            =   "frmDt041.frx":2038
         Spin            =   "frmDt041.frx":2082
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
         Index           =   12
         Left            =   3120
         TabIndex        =   54
         Top             =   1020
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "競売手数料"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2340
         TabIndex        =   53
         Top             =   2400
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   3180
         TabIndex        =   52
         Top             =   2400
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   4020
         TabIndex        =   51
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label lblDetailPname 
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
         Left            =   2040
         TabIndex        =   49
         Top             =   3120
         Width           =   1905
      End
      Begin VB.Label lblItime 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "yyyy/mm/dd hh:mm:ss"
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
         Left            =   1560
         TabIndex        =   48
         Top             =   3540
         Width           =   2385
      End
      Begin VB.Label lblDetailPcode 
         Alignment       =   1  '右揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "99"
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
         Left            =   1560
         TabIndex        =   47
         Top             =   3120
         Width           =   405
      End
   End
   Begin VB.Frame fraKey 
      Height          =   1095
      Left            =   60
      TabIndex        =   20
      Top             =   600
      Width           =   10575
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   2280
         Picture         =   "frmDt041.frx":20AA
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   555
      End
      Begin imText6Ctl.imText txtBcode 
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   180
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   635
         Caption         =   "frmDt041.frx":23B4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":2422
         Key             =   "frmDt041.frx":2440
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   5
         Left            =   60
         TabIndex        =   22
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "買主コード"
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
      Begin imText6Ctl.imText txtBname 
         Height          =   345
         Left            =   2880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frmDt041.frx":2474
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":24E2
         Key             =   "frmDt041.frx":2500
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
         MaxLength       =   80
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
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
      Begin imNumber6Ctl.imNumber imnNum 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmDt041.frx":2534
         Caption         =   "frmDt041.frx":2554
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDt041.frx":25C2
         Keys            =   "frmDt041.frx":25E0
         Spin            =   "frmDt041.frx":262A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##"
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
         Index           =   2
         Left            =   60
         TabIndex        =   39
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "精算回数"
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
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   10980
      TabIndex        =   0
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmDt041.frx":2652
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDt041.frx":26C0
      Key             =   "frmDt041.frx":26DE
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
      Left            =   11160
      TabIndex        =   19
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmDt041.frx":2712
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDt041.frx":2780
      Key             =   "frmDt041.frx":279E
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
Attribute VB_Name = "frmDt041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAIN_TABLE_NAME = "DT041"

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    On Error Resume Next
    txtBcode.SetFocus

End Sub

'目　的　　：
'条　件　　：レコード移動クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub cmdDataMove_Click(Index As Integer)

    Screen.MousePointer = vbHourglass

'    With m_clsAdoRecordCtl
'        Select Case Index
'            Case 0:
'                m_clsAdoRecordCtl.MoveFirst
'            Case 1:
'                If Trim(txtBcode.Text) = "" Then
'                    Screen.MousePointer = vbDefault
'                    Exit Sub
'                End If
'                .KeyValue = Array(CLng(txtBcode.Text))
'                m_clsAdoRecordCtl.MovePrevious
'            Case 2:
'                If Trim(txtBcode.Text) = "" Then
'                    Screen.MousePointer = vbDefault
'                    Exit Sub
'                End If
'                .KeyValue = Array(CLng(txtBcode.Text))
'                m_clsAdoRecordCtl.MoveNext
'            Case 3:
'                m_clsAdoRecordCtl.MoveLast
'            Case Else
'                Exit Sub
'        End Select
'        If FieldsSet(True, m_clsAdoRecordCtl.RecordSet) = False Then Exit Sub
'    End With
    
    Screen.MousePointer = vbDefault
    
End Sub

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error Resume Next

    If MsgBox("実行しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    If optSyori(0).Value = True Or optSyori(1).Value = True Then
        '入力チェック
        If DoValidationChecks() = False Then Exit Sub
        If DataUpdate() = False Then Exit Sub
    ElseIf optSyori(2).Value = True Then
        If DataDelete() = False Then Exit Sub
    End If
    
    'フィールドクリア
    Call FieldsClear(0)

    txtBcode.SetFocus

End Sub

'目　的　　：
'条　件　　：終了クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

'目　的　　：
'条　件　　：検索クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub cmdSearch_Click()

    Screen.MousePointer = vbHourglass
    frmDt041Search.Adodc1.ConnectionString = g_clsAdoSQL.Connection.ConnectionString
    frmDt041Search.Adodc1.Refresh
    Screen.MousePointer = vbDefault
    
    frmDt041Search.Show vbModal

End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
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
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "買主精算データ保守"

    '重複起動のチェック
    If App.PrevInstance = True Then
        Unload Me
        End
    End If
    
    '処理ボタン
    optSyori(0).Value = True
    optSyori(1).Value = False
    optSyori(2).Value = False
    optSyori(3).Value = False
    optSyori(4).Value = False
    
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
'作成年月日：２００２／０８／２７
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
'引　数　　：0：全画面 1:キー部 2:明細部
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        txtBcode.Text = ""
        txtBcode.Tag = ""
        txtBname.Text = ""
        imnNum.Value = 0
        
        imnTotal.Value = 0
        imnTax.Value = 0
        imnKeep.Value = 0
        imnGtotal.Value = 0
        chkRdiv.Value = 0
        optR(0).Value = True
        imnRdate_Year.Text = ""
        imnRdate_Month.Text = ""
        imnRdate_Day.Text = ""
        lblDetailPcode.Caption = ""
        lblDetailPname.Caption = ""
        lblItime.Caption = ""
        
        '201107
        imnBrate2.Value = 0
    ElseIf intKubun = 1 Then
        txtBcode.Text = ""
        txtBcode.Tag = ""
        txtBname.Text = ""
        imnNum.Value = 0
    ElseIf intKubun = 2 Then
        imnTotal.Value = 0
        imnTax.Value = 0
        imnKeep.Value = 0
        imnGtotal.Value = 0
        chkRdiv.Value = 0
        optR(0).Value = True
        imnRdate_Year.Text = ""
        imnRdate_Month.Text = ""
        imnRdate_Day.Text = ""
        lblDetailPcode.Caption = ""
        lblDetailPname.Caption = ""
        lblItime.Caption = ""
        
        '201107
        imnBrate2.Value = 0
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub
Private Sub imnBrate2_GotFocus()

    imnBrate2.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBrate2_LostFocus()

    imnBrate2.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnGtotal_GotFocus()

    imnGtotal.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnGtotal_LostFocus()

    imnGtotal.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnKeep_GotFocus()

    imnKeep.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnKeep_LostFocus()

    imnKeep.BackColor = FOCUS_NO_COLOR
    Call Calc_Total

End Sub

Private Sub imnNum_GotFocus()

    imnNum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnNum_LostFocus()

    imnNum.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnNum_Validate(Cancel As Boolean)

    On Error GoTo imnNum_Validate_Err

    If Trim(txtBcode.Text) = "" Then
        txtBcode.SetFocus
        DoEvents
        Call MsgBox("買主コードを入力してください。", vbOKOnly + vbCritical, "")
        Exit Sub
    End If

    If Trim(imnNum.Text) = "" Then
        Cancel = True
        Call MsgBox("精算回数を入力してください。", vbOKOnly + vbCritical, "")
        Exit Sub
    End If

    If optSyori(0).Value = True Then
        If FieldsSet(False) = True Then
            Cancel = True
            Call MsgBox("既にデータが存在します。", vbOKOnly + vbCritical, "")
            Exit Sub
        End If
    Else
        If FieldsSet(True) = False Then
            Cancel = True
            Call MsgBox("データが存在しません。", vbOKOnly + vbCritical, "")
            Exit Sub
        End If
    End If
    
    Exit Sub
    
imnNum_Validate_Err:

    Call MsgBox("精算回数フォーカス移動前エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imnNum_Validate_Err")
    
End Sub

Private Sub imnRdate_Day_GotFocus()

    imnRdate_Day.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Day_LostFocus()

    imnRdate_Day.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnRdate_Month_GotFocus()

    imnRdate_Month.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Month_LostFocus()

    imnRdate_Month.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRdate_Year_GotFocus()

    imnRdate_Year.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Year_LostFocus()

    imnRdate_Year.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnTax_GotFocus()

    imnTax.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnTax_LostFocus()

    imnTax.BackColor = FOCUS_NO_COLOR
    Call Calc_Total
    
End Sub

Private Sub imnTotal_GotFocus()

    imnTotal.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnTotal_LostFocus()

    imnTotal.BackColor = FOCUS_NO_COLOR
    Call Calc_Total

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    txtBcode.SetFocus

End Sub

'目　的　　：
'条　件　　：処理区分ボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub optSyori_Click(Index As Integer)

    Dim intIndex1 As Integer
    Dim adoRecordset1 As New ADODB.RecordSet
    Dim strSQL As String

    On Error GoTo optSyori_Click_Err

    '画面クリア
    Call FieldsClear(0)
    
    '背景色の変更
    For intIndex1 = 0 To 4
        If intIndex1 = Index Then
            optSyori(intIndex1).BackColor = BUTTON_ON
        Else
            optSyori(intIndex1).BackColor = BUTTON_OFF
        End If
    Next intIndex1
    
    Select Case Index
        Case 0: '新規
            Call FieldsControl(0, True)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
        Case 1: '変更
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
        Case 2: '削除
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
        Case 3: '印刷
            Call FieldsControl(0, False)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
            frmPrintDialog.Show vbModal
        Case 4: '外部出力
    End Select

    On Error Resume Next
    txtBcode.SetFocus
    DoEvents
    
    Exit Sub

optSyori_Click_Err:

    Call MsgBox("処理区分クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")

End Sub

Private Sub txtBcode_Change()

    If Trim(txtBcode.Text) = "" Then Exit Sub

    If txtBcode.Tag <> txtBcode.Text Then
        If optSyori(0).Value Or optSyori(1).Value Then
            fraMeisai.Enabled = True
            DoEvents
        End If
    End If

End Sub

Private Sub txtBcode_GotFocus()

    txtBcode.Tag = txtBcode.Text
    txtBcode.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtBcode_LostFocus()

    txtBcode.Tag = ""
    txtBcode.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtBcode_Validate(Cancel As Boolean)

    Dim strBuff As String
    
    On Error GoTo txtBcode_Validate_Err
    
    If Trim(txtBcode.Text) = "" Then Exit Sub
    If txtBcode.Tag = txtBcode.Text Then Exit Sub
    
    txtBname.Text = Global_Get_Bname(g_clsAdoSQL, txtBcode.Text, g_strOdate, strBuff)

    If Trim(imnNum.Text) = "" Then Exit Sub

    If optSyori(0).Value = True Then
        If FieldsSet(False) = True Then
            Cancel = True
            Call MsgBox("既にデータが存在します。", vbOKOnly + vbCritical, "")
            Exit Sub
        End If
    Else
        If FieldsSet(True) = False Then
            Cancel = True
            Call MsgBox("データが存在しません。", vbOKOnly + vbCritical, "")
            Exit Sub
        End If
    End If

txtBcode_Validate_Err:

    Call MsgBox("買主コードフォーカス移動前エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "txtBcode_Validate_Err")
  
End Sub

Private Sub txtBname_GotFocus()

    txtBname.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtBname_LostFocus()

    txtBname.BackColor = FOCUS_NO_COLOR

End Sub

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtBcode.Text) = "" Then
        strErrMsg = "買主コードを入力してください。"
        txtBcode.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtBname.Text) = "" Then
        strErrMsg = "名称を入力してください。"
        txtBname.SetFocus
        GoTo ErrorTrap:
    End If
    If imnNum.Value = 0 Then
        strErrMsg = "精算回数を入力してください。"
        imnNum.SetFocus
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

'目　的　　：フィールドの制御
'条　件　　：
'結　果　　：
'引　数　　：intKbn 0:キー部 1:レコード移動 2:明細
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
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
    End Select
        
    Exit Sub

FieldsControl_Err:

    Call MsgBox("フィールドの制御エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsControl_Err")

End Sub

'目　的　　：フィールドのセット
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Public Function FieldsSet(blnVisible As Boolean, Optional adoRecordsetArg As Variant) As Boolean

    Dim adoRecordset1 As New ADODB.RecordSet
    Dim strSQL As String

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    If IsMissing(adoRecordsetArg) = False Then
        Set adoRecordset1 = adoRecordsetArg
    Else
        strSQL = "SELECT * FROM " & MAIN_TABLE_NAME & _
                 " WHERE Odate = '" & g_strOdate & "'" & _
                 " AND Bcode = " & txtBcode.Text & _
                 " AND Num = " & imnNum.Value
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
        txtBcode.Text = .Fields("Bcode")
        txtBname.Text = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
        imnNum.Value = .Fields("Num")
        imnTotal.Value = IIf(IsNull(.Fields("Total")), 0, .Fields("Total"))
        imnTax.Value = IIf(IsNull(.Fields("Tax")), 0, .Fields("Tax"))
        imnKeep.Value = IIf(IsNull(.Fields("Keep")), 0, .Fields("Keep"))
        imnGtotal.Value = IIf(IsNull(.Fields("Gtotal")), 0, .Fields("Gtotal"))
        chkRdiv.Value = IIf(IsNull(.Fields("Rdiv")), 0, .Fields("Rdiv"))
        If Not IsNull(.Fields("R")) Then
            If .Fields("R") = PAYMENT_DIV_CASH Then
                optR(0).Value = True
            ElseIf .Fields("R") = PAYMENT_DIV_CHECK Then
                optR(1).Value = True
            ElseIf .Fields("R") = PAYMENT_DIV_TRANSFER Then
                optR(2).Value = True
            End If
        Else
            optR(0).Value = True
        End If
        If Not IsNull(.Fields("Rdate")) Then
            imnRdate_Year.Text = left$(.Fields("Rdate"), 4)
            imnRdate_Month.Text = Mid$(.Fields("Rdate"), 6, 2)
            imnRdate_Day.Text = right$(.Fields("Rdate"), 2)
        Else
            imnRdate_Year.Text = ""
            imnRdate_Month.Text = ""
            imnRdate_Day.Text = ""
        End If
        If Not IsNull(.Fields("Itime")) Then
            lblItime.Caption = Trim(.Fields("Itime"))
        Else
            lblItime.Caption = ""
        End If
        If Not IsNull(.Fields("Pcode")) Then
            lblDetailPcode.Caption = Trim(.Fields("Pcode"))
lblDetailPname.Caption = ""
        Else
            lblDetailPcode.Caption = ""
            lblDetailPname.Caption = ""
        End If
        
        '201107
        imnBrate2.Value = IIf(IsNull(.Fields("Brate2")), 0, .Fields("Brate2"))
        
    End With
    
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
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
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.RecordSet

    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
    With adoRecordset1
        strSQL = "SELECT * FROM " & MAIN_TABLE_NAME & _
                 " WHERE Odate = '" & g_strOdate & "'" & _
                 " AND Bcode = " & txtBcode.Text & _
                 " AND Num = " & imnNum.Value
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        .Fields("Odate") = g_strOdate
        .Fields("Bcode") = txtBcode.Text
        .Fields("Bname") = txtBname.Text
        .Fields("Num") = imnNum.Value
        .Fields("Total") = imnTotal.Value
        .Fields("Tax") = imnTax.Value
        .Fields("Keep") = imnKeep.Value
        .Fields("Gtotal") = imnGtotal.Value
        If chkRdiv.Value = 0 Then
            .Fields("Rdiv") = 0
            .Fields("R") = Null
            .Fields("Rdate") = Null
        Else
            .Fields("Rdiv") = chkRdiv.Value
            If optR(0).Value = True Then
                .Fields("R") = PAYMENT_DIV_CASH
            ElseIf optR(1).Value = True Then
                 .Fields("R") = PAYMENT_DIV_CHECK
            ElseIf optR(2).Value = True Then
                 .Fields("R") = PAYMENT_DIV_TRANSFER
            End If
            If Trim(imnRdate_Year.Text) <> "" And Trim(imnRdate_Month.Text) <> "" And Trim(imnRdate_Day.Text) <> "" Then
                .Fields("Rdate") = imnRdate_Year.Text & "/" & Format(imnRdate_Month.Text, "00") & "/" & Format(imnRdate_Day.Text, "00")
            End If
        End If
        .Fields("Itime") = Format(Now(), "hh:mm:ss")
        .Fields("Pcode") = g_strPcode
        '201107
        .Fields("Brate2") = imnBrate2.Value
        
        .Update
    End With
    
    g_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    adoRecordset1.Close
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

'目　的　　：データの削除
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Function DataDelete() As Boolean

    Dim strSQL As String

    On Error GoTo DataDelete_Err
    
    If Trim(txtBcode.Text) = "" Then
        DataDelete = True
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    With g_clsAdoSQL.Connection
        .BeginTrans
        
        'データ削除
        strSQL = "DELETE FROM " & MAIN_TABLE_NAME & _
                 " WHERE Odate = '" & g_strOdate & "'" & _
                 " AND Bcode = " & txtBcode.Text & _
                 " AND Num = " & imnNum.Value
        .Execute strSQL
    
        .CommitTrans
    End With
    
    Screen.MousePointer = vbDefault
    
    DataDelete = True
    
    Exit Function

DataDelete_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    DataDelete = False
    Screen.MousePointer = vbDefault
    Call MsgBox("データの削除エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataDelete_Err")

End Function

'目　的　　：合計の計算
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２７
'更新履歴　：
'
Private Sub Calc_Total()

    On Error GoTo Calc_Total_Err
    
    '201107
    'imnGtotal.Value = CCur(imnTotal.Value) + CCur(imnTax.Value) + CCur(imnKeep.Value)
    imnGtotal.Value = CCur(imnTotal.Value) + CCur(imnTax.Value) + CCur(imnKeep.Value) + CCur(imnBrate2.Value)
    
    Exit Sub
    
Calc_Total_Err:

    Call MsgBox("合計の計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Total_Err")

End Sub


