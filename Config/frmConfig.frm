VERSION 5.00
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   8820
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8475
   Icon            =   "frmConfig.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   8475
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame4 
      Caption         =   "バックアップ情報"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "バックアップドライブ"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "ホスト情報"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "ダウンロードパス"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "クライアント情報"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "ログパス"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "実行ファイルパス"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "ワークＤＢパス"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "ワークＤＢ名"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "SQLサーバ接続情報"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "実行タイムアウト"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "パスワード"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "ユーザー名"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "データソース名"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "データベース名"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "サーバー名"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Align           =   2  '下揃え
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
         Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "登録(F12)"
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
'   プログラム名：クライアント環境設定
'   処理内容　　：
'   前提条件　　：
'   作成者　　　：株式会社 コム・エンジニアリング　渥美
'   作成年月日　：２００２／０５／２４
'   更新履歴　　：
'
'******************************************************************

Private m_clsReg As New clsReg          'レジストリ操作クラス

Private Declare Function GetLogicalDrives Lib "KERNEL32" () As Long

Private Sub cboBackupDrive_DropDown()

    Dim Drives As Long
    Dim I As Long
    Dim Bit As Long

    On Error GoTo cboBackupDrive_DropDown_Err

    Bit = 1
    cboBackupDrive.Clear

    'ドライブ一覧取得
    Drives = GetLogicalDrives()

    'メッセージ作成
    For I = Asc("A") To Asc("Z")
        If (Drives And Bit) <> 0 Then
            cboBackupDrive.AddItem Chr(I) & ":"
        End If
        Bit = Bit * 2
    Next

    Exit Sub

cboBackupDrive_DropDown_Err:

    Call MsgBox("ドロップダウン時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBackupDrive_DropDown_Err")

End Sub

Private Sub cboBackupDrive_GotFocus()

    cboBackupDrive.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub cboBackupDrive_LostFocus()

    cboBackupDrive.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：登録ボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If MsgBox("登録しますか？", vbQuestion + vbYesNo, "確認") = vbNo Then Exit Sub

    '入力チェック
    If DoValidationChecks() = False Then Exit Sub

    'データ登録
    If DataUpdate() = False Then Exit Sub
    
    '終了
    Unload Me
    End
    
    Exit Sub

cmdExecute_Click_Err:

    Call MsgBox("登録ボタンクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

'目　的　　：
'条　件　　：登録ボタンキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub cmdExecute_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：終了ボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

'目　的　　：
'条　件　　：終了ボタンキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub cmdExit_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：フォルダ選択ボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub cmdFile_Click(Index As Integer)
    
    Dim strPath As String
    
    On Error GoTo cmdFile_Click_Err
    
    'フォルダ選択ダイアログ表示
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

    Call MsgBox("フォルダ選択ボタンクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdFile_Click_Err")

End Sub

'目　的　　：
'条　件　　：フォームロード時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "環境設定"

    '重複起動のチェック
    If App.PrevInstance = True Then
        Unload Me
        End
    End If

    '画面クリア
    Call FildsClear
    
    'データ表示
    Call FildsSet
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    End

End Sub

'目　的　　：
'条　件　　：実行タイムアウトフォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imnQueryTimeout_GotFocus()

    imnQueryTimeout.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：実行タイムアウトキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imnQueryTimeout_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：実行タイムアウトフォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imnQueryTimeout_LostFocus()

    imnQueryTimeout.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：実行ファイルパスフォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtBin_GotFocus()

    imtBin.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：実行ファイルパスキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtBin_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：実行ファイルパスフォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtBin_LostFocus()

    imtBin.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：ダウンロードパスフォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtDownLoad_GotFocus()
    
    imtDownLoad.BackColor = FOCUS_STOP_COLOR

End Sub

'目　的　　：
'条　件　　：ダウンロードパスキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtDownLoad_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：ダウンロードパスフォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtDownLoad_LostFocus()
   
    imtDownLoad.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：データソース名フォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtDSN_GotFocus()

    imtDSN.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：データソース名キーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtDSN_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：データソース名フォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtDSN_LostFocus()

    imtDSN.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：フォーカスコントロール
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus
    
End Sub

'目　的　　：
'条　件　　：フォーカスコントロール
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtFocusFirst_GotFocus()

    imtBin.SetFocus

End Sub

'目　的　　：キーダウン時の処理
'条　件　　：各コントロールのキーダウン時に設定
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Sub KeyDownControl(KeyCode As Integer, Shift As Integer)

    On Error GoTo KeyDownControl_Err
    
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

    Call MsgBox("キーダウン時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "KeyDownControl_Err")

End Sub

'目　的　　：画面クリア
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
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

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FildsClear_Err")

End Sub

'目　的　　：データ表示
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function FildsSet() As Boolean

    On Error GoTo FildsSet_Err

    FildsSet = False

    'レジストリ読み込み
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
    Call MsgBox("データ表示エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FildsSet_Err")

End Function

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
    
    On Error GoTo DoValidationChecks_Err

    '実行ファイルパス
    If Trim(imtBin.Text) = "" Then
        imtBin.SetFocus
        strErrMsg = "実行ファイルパスを入力してください！！"
        GoTo ErrorTrap:
    End If
    'ワークＤＢパス
    If Trim(imtLDatabase.Text) = "" Then
        imtLDatabase.SetFocus
        strErrMsg = "ワークＤＢパスを入力してください！！"
        GoTo ErrorTrap:
    End If
    'ワークＤＢ名
    If Trim(imtLDBName.Text) = "" Then
        imtLDBName.SetFocus
        strErrMsg = "ワークＤＢ名を入力してください！！"
        GoTo ErrorTrap:
    End If
    'ログパス
    If Trim(imtLog.Text) = "" Then
        imtLog.SetFocus
        strErrMsg = "ログパスを入力してください！！"
        GoTo ErrorTrap:
    End If
    'サーバー名
    If Trim(imtServer.Text) = "" Then
        imtServer.SetFocus
        strErrMsg = "サーバー名を入力してください！！"
        GoTo ErrorTrap:
    End If
    'データベース名
    If Trim(imtSQLDBName.Text) = "" Then
        imtSQLDBName.SetFocus
        strErrMsg = "データベース名を入力してください！！"
        GoTo ErrorTrap:
    End If
    'ユーザー名
    If Trim(imtUID.Text) = "" Then
        imtUID.SetFocus
        strErrMsg = "ユーザー名を入力してください！！"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks = False
    Call MsgBox(strErrMsg, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function

'目　的　　：データ登録
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    On Error GoTo DataUpdate_Err

    'マウスポインタを砂時計に変更
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
    
    'マウスポインタを元に戻す
    Screen.MousePointer = vbDefault
    
    DataUpdate = True
    
    Exit Function

DataUpdate_Err:

    'マウスポインタを元に戻す
    Screen.MousePointer = vbDefault
    DataUpdate = False
    Call MsgBox("データ登録エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_Err")

End Function

'目　的　　：
'条　件　　：ワークＤＢパスフォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLDatabase_GotFocus()

    imtLDatabase.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：ワークＤＢパスキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLDatabase_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：ワークＤＢパスフォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLDatabase_LostFocus()

    imtLDatabase.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：ワークＤＢ名フォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLDBName_GotFocus()
    
    imtLDBName.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：ワークＤＢ名キーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLDBName_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：ワークＤＢ名フォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLDBName_LostFocus()

    imtLDBName.BackColor = FOCUS_NO_COLOR

End Sub

'目　的　　：
'条　件　　：ログパスフォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLog_GotFocus()

    imtLog.BackColor = FOCUS_STOP_COLOR

End Sub

'目　的　　：
'条　件　　：ログパスキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLog_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：ログパスフォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtLog_LostFocus()

    imtLog.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：パスワードフォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtPWD_GotFocus()

    imtPWD.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：パスワードキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtPWD_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：パスワードフォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtPWD_LostFocus()

    imtPWD.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：サーバー名フォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtServer_GotFocus()

    imtServer.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：サーバー名キーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtServer_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：サーバー名フォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtServer_LostFocus()

    imtServer.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：データベース名フォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtSQLDBName_GotFocus()

    imtSQLDBName.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：データベース名キーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtSQLDBName_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：データベース名フォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtSQLDBName_LostFocus()

    imtSQLDBName.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：
'条　件　　：ユーザー名フォーカス取得時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtUID_GotFocus()

    imtUID.BackColor = FOCUS_STOP_COLOR
    
End Sub

'目　的　　：
'条　件　　：ユーザー名キーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtUID_KeyDown(KeyCode As Integer, Shift As Integer)

    Call KeyDownControl(KeyCode, Shift)
    
End Sub

'目　的　　：
'条　件　　：ユーザー名フォーカス喪失時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub imtUID_LostFocus()

    imtUID.BackColor = FOCUS_NO_COLOR
    
End Sub
