VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmMt070 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   8265
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMt070.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10110
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraRecordSelector 
      BackColor       =   &H80000001&
      Height          =   615
      Left            =   7740
      TabIndex        =   34
      Top             =   0
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
         Picture         =   "frmMt070.frx":0CFA
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   38
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
         Picture         =   "frmMt070.frx":0E44
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   37
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
         Picture         =   "frmMt070.frx":0F8E
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   36
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
         Picture         =   "frmMt070.frx":10D8
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
   End
   Begin VB.Frame fra 
      BackColor       =   &H80000001&
      Height          =   735
      Left            =   60
      TabIndex        =   33
      Top             =   7500
      Width           =   9975
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         Height          =   495
         Left            =   60
         TabIndex        =   18
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
         Picture         =   "frmMt070.frx":1222
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8220
         TabIndex        =   20
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
         Picture         =   "frmMt070.frx":123E
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   6480
         TabIndex        =   19
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
         Picture         =   "frmMt070.frx":1398
      End
   End
   Begin VB.Frame fraSyori 
      BackColor       =   &H80000001&
      Height          =   615
      Left            =   60
      TabIndex        =   25
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
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
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
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   29
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
      BackColor       =   &H80000001&
      Height          =   6255
      Left            =   60
      TabIndex        =   23
      Top             =   1260
      Width           =   9975
      Begin imDate6Ctl.imDate imdAdddate 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   5760
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   661
         Calendar        =   "frmMt070.frx":17EA
         Caption         =   "frmMt070.frx":196A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":19D8
         Keys            =   "frmMt070.frx":19F6
         Spin            =   "frmMt070.frx":1A54
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "yyyy/mm/dd"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "yyyy/mm/dd"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "2005/09/03"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   38598
         CenturyMode     =   0
      End
      Begin imText6Ctl.imText txtBname 
         Height          =   345
         Left            =   1560
         TabIndex        =   2
         Top             =   180
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frmMt070.frx":1A7C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":1AEA
         Key             =   "frmMt070.frx":1B08
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
      Begin imText6Ctl.imText txtBkana 
         Height          =   360
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   635
         Caption         =   "frmMt070.frx":1B3C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":1BAA
         Key             =   "frmMt070.frx":1BC8
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
         MaxLength       =   20
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWWWWWWW"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   6
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
         TabIndex        =   30
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "名　称"
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
         Index           =   4
         Left            =   60
         TabIndex        =   31
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "カナ名称"
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
         LabelWidth      =   56
         LabelHeight     =   25
         LabelLeft       =   20
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
         TabIndex        =   32
         Top             =   1920
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "住　所"
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
         Left            =   60
         TabIndex        =   45
         Top             =   2820
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "電話番号"
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
         Index           =   2
         Left            =   60
         TabIndex        =   46
         Top             =   4500
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "領収書名"
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
         Index           =   6
         Left            =   60
         TabIndex        =   47
         Top             =   4920
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "領収書区分"
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
         TabIndex        =   48
         Top             =   5340
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "業務区分"
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
      Begin imText6Ctl.imText txtAddres 
         Height          =   825
         Left            =   1560
         TabIndex        =   6
         Top             =   1920
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   1455
         Caption         =   "frmMt070.frx":1BFC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":1C6A
         Key             =   "frmMt070.frx":1C88
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
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   80
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
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
      Begin imText6Ctl.imText txtTel 
         Height          =   360
         Left            =   1560
         TabIndex        =   7
         Top             =   2820
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   635
         Caption         =   "frmMt070.frx":1CBC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":1D2A
         Key             =   "frmMt070.frx":1D48
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
         MaxLength       =   15
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWW"
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
      Begin imText6Ctl.imText txtRname 
         Height          =   345
         Left            =   1560
         TabIndex        =   11
         Top             =   4500
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frmMt070.frx":1D7C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":1DEA
         Key             =   "frmMt070.frx":1E08
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'なし
         Height          =   555
         Left            =   1500
         TabIndex        =   49
         Top             =   4740
         Width           =   2955
         Begin VB.OptionButton optRdiv 
            Caption         =   "出力する"
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
            Left            =   60
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   12
            Top             =   180
            Width           =   1395
         End
         Begin VB.OptionButton optRdiv 
            Caption         =   "出力しない"
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
            Left            =   1500
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   13
            Top             =   180
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'なし
         Height          =   495
         Left            =   1560
         TabIndex        =   50
         Top             =   5220
         Width           =   4335
         Begin VB.OptionButton optFdiv 
            Caption         =   "出品者"
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
            TabIndex        =   15
            Top             =   120
            Width           =   1395
         End
         Begin VB.OptionButton optFdiv 
            Caption         =   "買主"
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
            TabIndex        =   14
            Top             =   120
            Width           =   1395
         End
         Begin VB.OptionButton optFdiv 
            Caption         =   "両　方"
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
            TabIndex        =   16
            Top             =   120
            Width           =   1395
         End
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   8
         Left            =   60
         TabIndex        =   51
         Top             =   1440
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "郵便番号"
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
      Begin imText6Ctl.imText txtPost 
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   635
         Caption         =   "frmMt070.frx":1E3C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":1EAA
         Key             =   "frmMt070.frx":1EC8
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   10
         Left            =   60
         TabIndex        =   52
         Top             =   3240
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ＦＡＸ"
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
         LabelWidth      =   30
         LabelHeight     =   25
         LabelLeft       =   33
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
      Begin imText6Ctl.imText txtFax 
         Height          =   360
         Left            =   1560
         TabIndex        =   8
         Top             =   3240
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   635
         Caption         =   "frmMt070.frx":1EFC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":1F6A
         Key             =   "frmMt070.frx":1F88
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
         MaxLength       =   15
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWW"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   11
         Left            =   60
         TabIndex        =   53
         Top             =   3660
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "携帯/自宅"
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
         LabelWidth      =   68
         LabelHeight     =   25
         LabelLeft       =   14
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
      Begin imText6Ctl.imText txtPtel 
         Height          =   360
         Left            =   1560
         TabIndex        =   9
         Top             =   3660
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   635
         Caption         =   "frmMt070.frx":1FBC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":202A
         Key             =   "frmMt070.frx":2048
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
         MaxLength       =   15
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWW"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   12
         Left            =   60
         TabIndex        =   54
         Top             =   4080
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "メール"
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
      Begin imText6Ctl.imText txtMail 
         Height          =   360
         Left            =   1560
         TabIndex        =   10
         Top             =   4080
         Width           =   5715
         _Version        =   65536
         _ExtentX        =   10081
         _ExtentY        =   635
         Caption         =   "frmMt070.frx":207C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":20EA
         Key             =   "frmMt070.frx":2108
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
         MaxLength       =   30
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   13
         Left            =   60
         TabIndex        =   55
         Top             =   1020
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "代表者名"
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
      Begin imText6Ctl.imText txtCeo 
         Height          =   345
         Left            =   1560
         TabIndex        =   4
         Top             =   1020
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frmMt070.frx":213C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":21AA
         Key             =   "frmMt070.frx":21C8
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   14
         Left            =   60
         TabIndex        =   56
         Top             =   5760
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "登録日"
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
   End
   Begin VB.Frame fraKey 
      BackColor       =   &H80000001&
      Height          =   615
      Left            =   60
      TabIndex        =   22
      Top             =   660
      Width           =   9975
      Begin VB.CommandButton cmdRireki 
         Caption         =   "購入履歴(F11)"
         Height          =   375
         Left            =   5280
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   180
         Width           =   1515
      End
      Begin VB.CommandButton cmdHenkou 
         Caption         =   "買主番号変更(F10)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CommandButton cmdMt071 
         Caption         =   "買主番号履歴(F10)"
         Height          =   375
         Left            =   7440
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkAutoCode 
         Caption         =   "ｺｰﾄﾞ自動採番"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   2340
         Picture         =   "frmMt070.frx":21FC
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   180
         Width           =   555
      End
      Begin imText6Ctl.imText txtBcode 
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   180
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   635
         Caption         =   "frmMt070.frx":2506
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt070.frx":2574
         Key             =   "frmMt070.frx":2592
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
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "99999"
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
         TabIndex        =   24
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "得意先コード"
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
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   10560
      TabIndex        =   0
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmMt070.frx":25C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt070.frx":2634
      Key             =   "frmMt070.frx":2652
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
      Left            =   10800
      TabIndex        =   21
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmMt070.frx":2686
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt070.frx":26F4
      Key             =   "frmMt070.frx":2712
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
Attribute VB_Name = "frmMt070"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_clsAdoSQL As New clsAdoCore
Public m_clsReg As New clsReg
Public m_clsAdoRecordCtl As New clsAdoRecordCtl

Const AUTO_CODE = 1

'目　的　　：
'条　件　　：ｺｰﾄﾞ自動採番クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub chkAutoCode_Click()

    On Error Resume Next

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then
         txtBcode.Text = AutoCodeSet
         If txtBcode.Enabled Then txtBcode.SetFocus
    End If

End Sub

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
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
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub cmdDataMove_Click(Index As Integer)

    Screen.MousePointer = vbHourglass

    With m_clsAdoRecordCtl
        Select Case Index
            Case 0:
                m_clsAdoRecordCtl.MoveFirst
            Case 1:
                If Trim(txtBcode.Text) = "" Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                .KeyValue = Array(CLng(txtBcode.Text))
                m_clsAdoRecordCtl.MovePrevious
            Case 2:
                If Trim(txtBcode.Text) = "" Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                .KeyValue = Array(CLng(txtBcode.Text))
                m_clsAdoRecordCtl.MoveNext
            Case 3:
                m_clsAdoRecordCtl.MoveLast
            Case Else
                Exit Sub
        End Select
        If FieldsSet(True, m_clsAdoRecordCtl.RecordSet) = False Then Exit Sub
    End With
    
    Screen.MousePointer = vbDefault
    
End Sub

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub cmdExecute_Click()

    Dim varAdddate As Variant

    On Error Resume Next

    If MsgBox("実行しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    If optSyori(0).Value = True Or optSyori(1).Value = True Then
        '入力チェック
        If DoValidationChecks() = False Then Exit Sub
        If DataUpdate() = False Then Exit Sub
    ElseIf optSyori(2).Value = True Then
        If DataDelete() = False Then Exit Sub
    End If
    
    If optSyori(0).Value = True Then
        varAdddate = imdAdddate.Number
    End If
    
    'フィールドクリア
    Call FieldsClear(0)

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then txtBcode.Text = AutoCodeSet
    
    If optSyori(0).Value = True Then
        imdAdddate.Number = varAdddate
    End If
    
    txtBcode.SetFocus

End Sub

'目　的　　：
'条　件　　：終了クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdHenkou_Click()

    frmMt072.Show vbModal

End Sub

Private Sub cmdMt071_Click()

    If Trim(txtBcode.Text) = "" Then
        DoEvents
        Call MsgBox("得意先コードを入力してください。", vbOKOnly + vbCritical, "")
        txtBcode.SetFocus
        DoEvents
        Exit Sub
    End If
    If optFdiv(1).Value = True Then
         DoEvents
        Call MsgBox("買主を指定してください。", vbOKOnly + vbCritical, "")
        txtBcode.SetFocus
        DoEvents
        Exit Sub
    End If
    frmMt071.Show vbModal

End Sub

Private Sub cmdRireki_Click()

    Call Shell(m_clsReg.Bin & "\" & "RYPMF170.exe" & " " & txtBcode.Text, vbNormalFocus)

End Sub

'目　的　　：
'条　件　　：検索クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub cmdSearch_Click()

    Screen.MousePointer = vbHourglass
    frmMt070Search.Adodc1.ConnectionString = m_clsAdoSQL.Connection.ConnectionString
    frmMt070Search.Adodc1.Refresh
    Screen.MousePointer = vbDefault
    
    frmMt070Search.Show vbModal

End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
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
'            If cmdMt071.Enabled = False Then Exit Sub
'            cmdMt071.SetFocus
'            DoEvents
'            Call cmdMt071_Click
            
            If cmdHenkou.Enabled = False Then Exit Sub
            cmdHenkou.SetFocus
            DoEvents
            Call cmdHenkou_Click
        Case vbKeyF11
            If cmdRireki.Enabled = False Then Exit Sub
            cmdRireki.SetFocus
            DoEvents
            Call cmdRireki_Click
        
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
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

'    Me.Caption = SYSTEM_NAME & "-" & "得意先マスタ保守"
    Me.Caption = "得意先マスタ保守(過去分参照)"

    '重複起動のチェック
    If App.PrevInstance = True Then
        Unload Me
        End
    End If
        
    'レジストリ読み込み
    m_clsReg.RegKey = REG_KEY
    If m_clsReg.ReadReg = False Then
        Unload Me
        End
    End If

    'データベース接続
    With m_clsAdoSQL
        .Provider = adoSQLServer
        .Server = m_clsReg.Server
'        .DBName = m_clsReg.DBName
        .DBName = "YAWASEDB_BK;Persist Security Info=True" '(過去分参照)
        .UID = m_clsReg.UID
        .PWD = m_clsReg.PWD
        .CommandTimeOut = m_clsReg.CommandTimeOut
        If .Connect = False Then
            Unload Me
            End
        End If
    End With
    
    'レコード移動
    With m_clsAdoRecordCtl
        .Connection = m_clsAdoSQL.Connection
        .TableName = "MT070"
        .KeyName = Array("Bcode")
    End With
    
    '処理ボタン
    optSyori(0).Value = True
    optSyori(1).Value = False
    optSyori(2).Value = False
    optSyori(3).Value = False
    optSyori(4).Value = False
    
'    chkAutoCode.Value = AUTO_CODE
    chkAutoCode.Value = 0
    If chkAutoCode.Value = 1 Then txtBcode.Text = AutoCodeSet
    
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
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set m_clsAdoSQL = Nothing
    Set m_clsReg = Nothing
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
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        txtBcode.Text = ""
        txtBcode.Tag = ""
        
        txtBname.Text = ""
        txtBname.Tag = ""
        txtBkana.Text = ""
        txtBkana.Tag = ""
        txtCeo.Text = ""
        txtPost.Text = ""
        txtAddres.Text = ""
        txtTel.Text = ""
        txtFax.Text = ""
        txtPtel.Text = ""
        txtMail.Text = ""
        txtRname.Text = ""
        optRdiv(0).Value = True
        optRdiv(1).Value = False
        optFdiv(0).Value = True
        optFdiv(1).Value = False
        optFdiv(2).Value = False
        imdAdddate.Value = Null
        
        cmdHenkou.Enabled = False
        cmdRireki.Enabled = False
        
    ElseIf intKubun = 1 Then
        txtBcode.Text = ""
        txtBcode.Tag = ""
    ElseIf intKubun = 2 Then
        txtBname.Text = ""
        txtBname.Tag = ""
        txtBkana.Text = ""
        txtBkana.Tag = ""
        txtCeo.Text = ""
        txtPost.Text = ""
        txtAddres.Text = ""
        txtTel.Text = ""
        txtFax.Text = ""
        txtPtel.Text = ""
        txtMail.Text = ""
        txtRname.Text = ""
        optRdiv(0).Value = True
        optRdiv(1).Value = False
        optFdiv(0).Value = True
        optFdiv(1).Value = False
        optFdiv(2).Value = False
        imdAdddate.Value = Null
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

Private Sub imdAdddate_GotFocus()

   imdAdddate.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub imdAdddate_LostFocus()

   imdAdddate.BackColor = FOCUS_NO_COLOR
   
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
'作成年月日：２００２／０６／１０
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
            If chkAutoCode.Value = 1 Then txtBcode.Text = AutoCodeSet
            imdAdddate.Number = Format(Now(), "yyyymmdd")
        Case 1: '変更
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
            cmdHenkou.Enabled = False
            cmdRireki.Enabled = True
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
            Call FieldsControl(0, False)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
            
            'Excel出力
            strSQL = "SELECT * FROM vw_MT070"
            adoRecordset1.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoRecordset1.EOF = True Then
                Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
                Exit Sub
            End If
            
            Dim objClsExcelOut As New clsExcelOut
            objClsExcelOut.TitleName = Array("得意先コード", "名称", "カナ名称", "郵便番号", "住所", "電話", "FAX", "携帯/自宅", "メール", "領収書名", "領収書区分", "業務区分", "代表者名")
            objClsExcelOut.RecordSet = adoRecordset1
            objClsExcelOut.OutPut
            Set objClsExcelOut = Nothing
    End Select

    On Error Resume Next
    txtBcode.SetFocus
    DoEvents
    
    Exit Sub

optSyori_Click_Err:

    Call MsgBox("処理区分クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")

End Sub

Private Sub txtAddres_GotFocus()

    txtAddres.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtAddres_LostFocus()

    txtAddres.BackColor = FOCUS_NO_COLOR

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

    If Trim(txtBcode.Text) = "" Then Exit Sub
    If txtBcode.Tag = txtBcode.Text Then Exit Sub

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

End Sub

Private Sub txtBname_Validate(Cancel As Boolean)

    If txtBname.Tag <> txtBname.Text And Trim(txtRname.Text) = "" Then
        txtRname.Text = txtBname.Text
    End If
    
End Sub

Private Sub txtBname_Furigana(Yomi As String)

    txtBkana.Text = txtBkana.Text & Yomi

End Sub

Private Sub txtBname_GotFocus()

    If Trim(txtBkana.Text) = "" Then
        txtBname.FuriganaOn = True
    Else
        txtBname.FuriganaOn = False
    End If
    txtBname.BackColor = FOCUS_STOP_COLOR
    txtBname.Tag = txtBname.Text

End Sub

Private Sub txtBname_LostFocus()

    txtBname.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtBkana_GotFocus()

    txtBkana.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtBkana_LostFocus()

    txtBkana.BackColor = FOCUS_NO_COLOR

End Sub

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtBcode.Text) = "" Then
        strErrMsg = "得意先コードを入力してください。"
        txtBcode.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtBname.Text) = "" Then
        strErrMsg = "名称を入力してください。"
        txtBname.SetFocus
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
'作成年月日：２００２／０６／１０
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
'作成年月日：２００２／０６／１０
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
        strSQL = "{call sp_MT070;2(" & txtBcode.Text & ")}"
        adoRecordset1.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
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
        txtBkana.Text = IIf(IsNull(.Fields("Bkana")), "", Trim(.Fields("Bkana")))
        txtCeo.Text = IIf(IsNull(.Fields("Ceo")), "", Trim(.Fields("Ceo")))
        txtPost.Text = IIf(IsNull(.Fields("Post")), "", Trim(.Fields("Post")))
        txtAddres.Text = IIf(IsNull(.Fields("Addres")), "", Trim(.Fields("Addres")))
        txtTel.Text = IIf(IsNull(.Fields("Tel")), "", Trim(.Fields("Tel")))
        txtFax.Text = IIf(IsNull(.Fields("Fax")), "", Trim(.Fields("Fax")))
        txtPtel.Text = IIf(IsNull(.Fields("Ptel")), "", Trim(.Fields("Ptel")))
        txtMail.Text = IIf(IsNull(.Fields("Mail")), "", Trim(.Fields("Mail")))
        txtRname.Text = IIf(IsNull(.Fields("Rname")), "", Trim(.Fields("Rname")))
        If Not IsNull(.Fields("Rdiv")) Then
            Select Case .Fields("Rdiv")
                Case RECEIPT_OFF:
                    optRdiv(1).Value = True
                Case RECEIPT_ON:
                    optRdiv(0).Value = True
            End Select
        End If
        If Not IsNull(.Fields("Fdiv")) Then
            Select Case .Fields("Fdiv")
                Case BUSINESS_DIV_EXHIBITION:
                    optFdiv(1).Value = True
                Case BUSINESS_DIV_BUYER:
                    optFdiv(0).Value = True
                Case BUSINESS_DIV_ALL:
                    optFdiv(2).Value = True
            End Select
        End If
        If IsNull(.Fields("Adddate")) = True Then
            imdAdddate.Number = 0
        Else
            imdAdddate.Number = Format(.Fields("Adddate"), "yyyymmdd")
        End If
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
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.RecordSet

    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    m_clsAdoSQL.Connection.BeginTrans
    
    If Trim(txtRname.Text) = "" Then txtRname.Text = Trim(txtBname.Text)
    
    With adoRecordset1
        strSQL = "{call sp_MT070;2(" & txtBcode.Text & ")}"
        .Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        .Fields("Bcode") = txtBcode.Text
        .Fields("Bname") = txtBname.Text
        .Fields("Bkana") = txtBkana.Text
        .Fields("Post") = txtPost.Text
        .Fields("Addres") = txtAddres.Text
        .Fields("Tel") = txtTel.Text
        .Fields("Fax") = txtFax.Text
        .Fields("Ptel") = txtPtel.Text
        .Fields("Mail") = txtMail.Text
        .Fields("Rname") = txtRname.Text
        If optRdiv(0).Value = True Then
            .Fields("Rdiv") = RECEIPT_ON
        Else
            .Fields("Rdiv") = RECEIPT_OFF
        End If
        If optFdiv(0).Value = True Then
            .Fields("Fdiv") = BUSINESS_DIV_BUYER
        ElseIf optFdiv(1).Value = True Then
            .Fields("Fdiv") = BUSINESS_DIV_EXHIBITION
        ElseIf optFdiv(2).Value = True Then
            .Fields("Fdiv") = BUSINESS_DIV_ALL
        End If
        .Fields("Ceo") = txtCeo.Text
        .Fields("Adddate") = imdAdddate.Value
        .Update
    End With
    
    m_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
    DataUpdate = True
    
    Exit Function

DataUpdate_Err:

    m_clsAdoSQL.Connection.RollbackTrans
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
'作成年月日：２００２／０６／１０
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
    
    With m_clsAdoSQL.Connection
        .BeginTrans
        
        'データ削除
        strSQL = "{call sp_MT070;9(" & txtBcode.Text & ")}"
        .Execute strSQL
    
        .CommitTrans
    End With
    
    Screen.MousePointer = vbDefault
    
    DataDelete = True
    
    Exit Function

DataDelete_Err:

    m_clsAdoSQL.Connection.RollbackTrans
    DataDelete = False
    Screen.MousePointer = vbDefault
    Call MsgBox("データの削除エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataDelete_Err")

End Function

'目　的　　：コードの自動採番
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Function AutoCodeSet() As String

    Dim adoRecordset1 As New ADODB.RecordSet
    Dim strSQL As String

    On Error GoTo AutoCodeSet_Err
    
    AutoCodeSet = ""
    
    Screen.MousePointer = vbHourglass
    
    With adoRecordset1
        strSQL = "{call sp_MT070;1}"
        .Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If .EOF Or .BOF Then
            AutoCodeSet = 1
            adoRecordset1.Close
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        .MoveLast
        If CLng(.Fields("Bcode")) < 99999 Then
            AutoCodeSet = CLng(.Fields("Bcode")) + 1
        End If
    End With
    
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function

AutoCodeSet_Err:

    AutoCodeSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("コードの自動採番エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "AutoCodeSet_Err")

End Function

Private Sub txtCeo_GotFocus()

   txtCeo.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtCeo_LostFocus()

   txtCeo.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtFax_GotFocus()

   txtFax.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtFax_LostFocus()

   txtFax.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtMail_GotFocus()

   txtMail.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtMail_LostFocus()

   txtMail.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtPost_GotFocus()

   txtPost.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtPost_LostFocus()

   txtPost.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtPtel_GotFocus()

   txtPtel.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtPtel_LostFocus()

   txtPtel.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtRname_GotFocus()

   txtRname.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtRname_LostFocus()

   txtRname.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtTel_GotFocus()
    
    txtTel.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtTel_LostFocus()
    
    txtTel.BackColor = FOCUS_NO_COLOR
    
End Sub
