VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmMt030 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   3060
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
   Icon            =   "frmMt030.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   10110
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraRecordSelector 
      Height          =   615
      Left            =   7740
      TabIndex        =   18
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
         Picture         =   "frmMt030.frx":0CFA
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   22
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
         Picture         =   "frmMt030.frx":0E44
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   21
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
         Picture         =   "frmMt030.frx":0F8E
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   20
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
         Picture         =   "frmMt030.frx":10D8
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
   End
   Begin VB.Frame fra 
      Height          =   735
      Left            =   60
      TabIndex        =   17
      Top             =   2280
      Width           =   9975
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         Height          =   495
         Left            =   60
         TabIndex        =   4
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
         Picture         =   "frmMt030.frx":1222
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   6
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
         Picture         =   "frmMt030.frx":123E
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   6420
         TabIndex        =   5
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
         Picture         =   "frmMt030.frx":1398
      End
   End
   Begin VB.Frame fraSyori 
      Height          =   615
      Left            =   60
      TabIndex        =   11
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   15
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
      Height          =   1095
      Left            =   60
      TabIndex        =   9
      Top             =   1200
      Width           =   9975
      Begin imText6Ctl.imText txtPname 
         Height          =   345
         Left            =   1560
         TabIndex        =   2
         Top             =   180
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   609
         Caption         =   "frmMt030.frx":17EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt030.frx":1858
         Key             =   "frmMt030.frx":1876
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
         TabIndex        =   16
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
         Index           =   0
         Left            =   60
         TabIndex        =   27
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "パスワード"
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
      Begin imText6Ctl.imText txtPwd 
         Height          =   345
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "frmMt030.frx":18BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt030.frx":1928
         Key             =   "frmMt030.frx":1946
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
         PasswordChar    =   "*"
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   5
         LengthAsByte    =   -1
         Text            =   "WWWWW"
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
   End
   Begin VB.Frame fraKey 
      Height          =   615
      Left            =   60
      TabIndex        =   8
      Top             =   600
      Width           =   9975
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   2100
         Picture         =   "frmMt030.frx":198A
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   555
      End
      Begin imText6Ctl.imText txtPcode 
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   180
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   635
         Caption         =   "frmMt030.frx":1C94
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt030.frx":1D02
         Key             =   "frmMt030.frx":1D20
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
         Index           =   5
         Left            =   60
         TabIndex        =   10
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "担当者ｺｰﾄﾞ"
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
         LabelWidth      =   72
         LabelHeight     =   25
         LabelLeft       =   12
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
      Caption         =   "frmMt030.frx":1D64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt030.frx":1DD2
      Key             =   "frmMt030.frx":1DF0
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
      TabIndex        =   7
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmMt030.frx":1E34
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt030.frx":1EA2
      Key             =   "frmMt030.frx":1EC0
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
Attribute VB_Name = "frmMt030"
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
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Sub chkAutoCode_Click()

    On Error Resume Next

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then
         txtPcode.Text = AutoCodeSet
         If txtPcode.Enabled Then txtPcode.SetFocus
    End If

End Sub

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    On Error Resume Next
    txtPcode.SetFocus

End Sub

'目　的　　：
'条　件　　：レコード移動クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Sub cmdDataMove_Click(Index As Integer)

    Screen.MousePointer = vbHourglass

    With m_clsAdoRecordCtl
        Select Case Index
            Case 0:
                m_clsAdoRecordCtl.MoveFirst
            Case 1:
                If Trim(txtPcode.Text) = "" Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                .KeyValue = Array(CLng(txtPcode.Text))
                m_clsAdoRecordCtl.MovePrevious
            Case 2:
                If Trim(txtPcode.Text) = "" Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                .KeyValue = Array(CLng(txtPcode.Text))
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
'作成年月日：２００２／０６／１１
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

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then txtPcode.Text = AutoCodeSet

    txtPcode.SetFocus

End Sub

'目　的　　：
'条　件　　：終了クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１１
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
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Sub cmdSearch_Click()

    Screen.MousePointer = vbHourglass
    frmMt030Search.Adodc1.ConnectionString = m_clsAdoSQL.Connection.ConnectionString
    frmMt030Search.Adodc1.Refresh
    Screen.MousePointer = vbDefault
    
    frmMt030Search.Show vbModal

End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１１
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
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "担当者マスタ保守"

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
        .DBName = m_clsReg.DBName
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
        .TableName = "MT030"
        .KeyName = Array("Pcode")
    End With
    
    '処理ボタン
    optSyori(0).Value = True
    optSyori(1).Value = False
    optSyori(2).Value = False
    optSyori(3).Value = False
    optSyori(4).Value = False
    
    chkAutoCode.Value = AUTO_CODE
    If chkAutoCode.Value = 1 Then txtPcode.Text = AutoCodeSet
    
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
'作成年月日：２００２／０６／１１
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
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        txtPcode.Text = ""
        txtPcode.Tag = ""
        
        txtPname.Text = ""
        txtPname.Tag = ""
        txtPwd.Text = ""
    ElseIf intKubun = 1 Then
        txtPcode.Text = ""
        txtPcode.Tag = ""
    ElseIf intKubun = 2 Then
        txtPname.Text = ""
        txtPname.Tag = ""
        txtPwd.Text = ""
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    txtPcode.SetFocus

End Sub

'目　的　　：
'条　件　　：処理区分ボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１１
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
            If chkAutoCode.Value = 1 Then txtPcode.Text = AutoCodeSet
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
            Call FieldsControl(0, False)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
            
            'Excel出力
            strSQL = "SELECT * FROM vw_MT030"
            adoRecordset1.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoRecordset1.EOF = True Then
                Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
                Exit Sub
            End If
            
            Dim objClsExcelOut As New clsExcelOut
            objClsExcelOut.TitleName = Array("担当者コード", "名称", "パスワード")
            objClsExcelOut.RecordSet = adoRecordset1
            objClsExcelOut.OutPut
            Set objClsExcelOut = Nothing
    End Select

    On Error Resume Next
    txtPcode.SetFocus
    DoEvents
    
    Exit Sub

optSyori_Click_Err:

    Call MsgBox("処理区分クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")

End Sub

Private Sub txtPcode_Change()

    If Trim(txtPcode.Text) = "" Then Exit Sub

    If txtPcode.Tag <> txtPcode.Text Then
        If optSyori(0).Value Or optSyori(1).Value Then
            fraMeisai.Enabled = True
            DoEvents
        End If
    End If

End Sub

Private Sub txtPcode_GotFocus()

    txtPcode.Tag = txtPcode.Text
    txtPcode.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtPcode_LostFocus()

    txtPcode.Tag = ""
    txtPcode.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtPcode_Validate(Cancel As Boolean)

    If Trim(txtPcode.Text) = "" Then Exit Sub
    If txtPcode.Tag = txtPcode.Text Then Exit Sub

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

Private Sub txtPname_GotFocus()

    txtPname.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtPname_LostFocus()

    txtPname.BackColor = FOCUS_NO_COLOR

End Sub

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtPcode.Text) = "" Then
        strErrMsg = "担当者コードを入力してください。"
        txtPcode.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtPname.Text) = "" Then
        strErrMsg = "名称を入力してください。"
        txtPname.SetFocus
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
'作成年月日：２００２／０６／１１
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
'作成年月日：２００２／０６／１１
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
        strSQL = "{call sp_MT030;2(" & txtPcode.Text & ")}"
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
        txtPcode.Text = .Fields("Pcode")
        txtPname.Text = IIf(IsNull(.Fields("Pname")), "", Trim(.Fields("Pname")))
        txtPwd.Text = IIf(IsNull(.Fields("Pwd")), "", Trim(.Fields("Pwd")))
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
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.RecordSet

    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    m_clsAdoSQL.Connection.BeginTrans
    
    With adoRecordset1
        strSQL = "{call sp_MT030;2(" & txtPcode.Text & ")}"
        .Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        .Fields("Pcode") = txtPcode.Text
        .Fields("Pname") = txtPname.Text
        .Fields("Pwd") = txtPwd.Text
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
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Function DataDelete() As Boolean

    Dim strSQL As String

    On Error GoTo DataDelete_Err
    
    If Trim(txtPcode.Text) = "" Then
        DataDelete = True
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    With m_clsAdoSQL.Connection
        .BeginTrans
        
        'データ削除
        strSQL = "{call sp_MT030;9(" & txtPcode.Text & ")}"
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
'作成年月日：２００２／０６／１１
'更新履歴　：
'
Private Function AutoCodeSet() As String

    Dim adoRecordset1 As New ADODB.RecordSet
    Dim strSQL As String

    On Error GoTo AutoCodeSet_Err
    
    AutoCodeSet = ""
    
    Screen.MousePointer = vbHourglass
    
    With adoRecordset1
        strSQL = "{call sp_MT030;1}"
        .Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If .EOF Or .BOF Then
            AutoCodeSet = 1
            adoRecordset1.Close
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        .MoveLast
        If CLng(.Fields("Pcode")) < 99 Then
            AutoCodeSet = CLng(.Fields("Pcode")) + 1
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

Private Sub txtPwd_GotFocus()

    txtPwd.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtPwd_LostFocus()

    txtPwd.BackColor = FOCUS_NO_COLOR

End Sub
