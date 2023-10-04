VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmYpmf190 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   7860
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf190.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   13740
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   12
      Top             =   7020
      Width           =   13515
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   7
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
         Picture         =   "frmYpmf190.frx":0CFA
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   11640
         TabIndex        =   9
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
         Picture         =   "frmYpmf190.frx":0D16
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   9900
         TabIndex        =   8
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
         Picture         =   "frmYpmf190.frx":0E70
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   13800
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf190.frx":12C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf190.frx":1330
      Key             =   "frmYpmf190.frx":134E
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
      Left            =   13860
      TabIndex        =   10
      Top             =   360
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf190.frx":1392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf190.frx":1400
      Key             =   "frmYpmf190.frx":141E
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
      TabIndex        =   11
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf190.frx":1462
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf190.frx":14D0
      Key             =   "frmYpmf190.frx":14EE
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
   Begin VB.Frame fraUketuke 
      Height          =   6255
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   13275
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "対象年月日"
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
         Caption         =   "frmYpmf190.frx":1532
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":15A0
         Key             =   "frmYpmf190.frx":15BE
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
         Caption         =   "frmYpmf190.frx":15F2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1660
         Key             =   "frmYpmf190.frx":167E
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
         Caption         =   "frmYpmf190.frx":16B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1720
         Key             =   "frmYpmf190.frx":173E
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
         Caption         =   "frmYpmf190.frx":1772
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":17E0
         Key             =   "frmYpmf190.frx":17FE
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
         Caption         =   "frmYpmf190.frx":1832
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":18A0
         Key             =   "frmYpmf190.frx":18BE
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
         Caption         =   "frmYpmf190.frx":18F2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1960
         Key             =   "frmYpmf190.frx":197E
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
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   22
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   21
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   20
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   19
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   18
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   17
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label4 
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
         Left            =   5340
         TabIndex        =   16
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame fraJpg 
      Height          =   6315
      Left            =   240
      TabIndex        =   23
      Top             =   540
      Width           =   13275
      Begin VB.CommandButton cmdDel 
         Caption         =   "一覧から　削除"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   660
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1035
      End
      Begin VB.CommandButton cmdFileClear 
         Caption         =   "やり直し"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   660
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   4140
         Width           =   1035
      End
      Begin VB.ListBox List1 
         Height          =   3000
         Left            =   1800
         TabIndex        =   50
         Top             =   3060
         Width           =   4155
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   180
         Top             =   5580
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.PictureBox Picture1 
         Height          =   795
         Left            =   780
         ScaleHeight     =   65.684
         ScaleMode       =   0  'ﾕｰｻﾞｰ
         ScaleWidth      =   63.072
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5340
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.FileListBox File1 
         Height          =   3030
         Left            =   1800
         MultiSelect     =   2  '拡張
         Pattern         =   "*.jpg;*.jpeg;*.bmp"
         TabIndex        =   34
         Top             =   3060
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton cmdSelectFolder 
         Caption         =   "フォルダを選ぶ"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   660
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1035
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   3060
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "写真の一覧"
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
         Index           =   2
         Left            =   240
         TabIndex        =   37
         Top             =   300
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "開催日"
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
         Left            =   240
         TabIndex        =   38
         Top             =   780
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "情報"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   39
         Top             =   1560
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "撮影日"
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
      Begin imText6Ctl.imText txtYear 
         Height          =   420
         Left            =   1800
         TabIndex        =   25
         Top             =   300
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   741
         Caption         =   "frmYpmf190.frx":19B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1A20
         Key             =   "frmYpmf190.frx":1A3E
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
         Left            =   3240
         TabIndex        =   26
         Top             =   300
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf190.frx":1A72
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1AE0
         Key             =   "frmYpmf190.frx":1AFE
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
         Left            =   4440
         TabIndex        =   27
         Top             =   300
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf190.frx":1B32
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1BA0
         Key             =   "frmYpmf190.frx":1BBE
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
      Begin imText6Ctl.imText txtSubtitle 
         Height          =   405
         Left            =   1800
         TabIndex        =   28
         Top             =   780
         Width           =   5115
         _Version        =   65536
         _ExtentX        =   9022
         _ExtentY        =   714
         Caption         =   "frmYpmf190.frx":1BF2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1C60
         Key             =   "frmYpmf190.frx":1C7E
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
      Begin imText6Ctl.imText txtTYear 
         Height          =   420
         Left            =   1800
         TabIndex        =   29
         Top             =   1560
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   741
         Caption         =   "frmYpmf190.frx":1CB2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1D20
         Key             =   "frmYpmf190.frx":1D3E
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
      Begin imText6Ctl.imText txtTMonth 
         Height          =   420
         Left            =   3240
         TabIndex        =   30
         Top             =   1560
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf190.frx":1D72
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1DE0
         Key             =   "frmYpmf190.frx":1DFE
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
      Begin imText6Ctl.imText txtTDay 
         Height          =   420
         Left            =   4440
         TabIndex        =   31
         Top             =   1560
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf190.frx":1E32
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1EA0
         Key             =   "frmYpmf190.frx":1EBE
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
      Begin imText6Ctl.imText txtTHour 
         Height          =   420
         Left            =   1800
         TabIndex        =   32
         Top             =   2040
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf190.frx":1EF2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":1F60
         Key             =   "frmYpmf190.frx":1F7E
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
      Begin imText6Ctl.imText txtTmin 
         Height          =   420
         Left            =   2880
         TabIndex        =   33
         Top             =   2040
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf190.frx":1FB2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf190.frx":2020
         Key             =   "frmYpmf190.frx":203E
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
      Begin VB.Image Image1 
         BorderStyle     =   1  '実線
         Height          =   4755
         Left            =   6180
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "(午後4時は「16」と入力)"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   1800
         TabIndex        =   49
         Top             =   2520
         Width           =   2595
      End
      Begin VB.Label Label1 
         Caption         =   "分頃"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   3540
         TabIndex        =   48
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "時"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   2400
         TabIndex        =   47
         Top             =   2100
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   5040
         TabIndex        =   46
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   45
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3900
         TabIndex        =   44
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "(特別大市、石材・資材など)"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   43
         Top             =   1200
         Width           =   4035
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   42
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   41
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3900
         TabIndex        =   40
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.ListBox List2 
      Height          =   3000
      Left            =   6240
      TabIndex        =   51
      Top             =   3600
      Visible         =   0   'False
      Width           =   5595
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6915
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   12197
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "受付データ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "写　真"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmYpmf190"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)

End Sub

Private Sub cmdDel_Click()

    If List1.ListIndex < 0 Then
        MsgBox "画像を選択してください", vbExclamation + vbOKOnly, ""
        Exit Sub
    End If
    
    List2.RemoveItem List1.ListIndex
    List1.RemoveItem List1.ListIndex
    List2.Refresh
    List1.Refresh

    Picture1.Picture = Nothing
    Image1.Picture = Nothing

End Sub

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err
    
    If MsgBox("実行しますか？", vbQuestion + vbYesNo, "") = vbNo Then Exit Sub
    
    If TabStrip1.Tabs(1).Selected = True Then
        '入力チェック
        If DoValidationChecks1() = False Then Exit Sub
        'データ転送
        If UploadData1() = False Then Exit Sub
    
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        '入力チェック
        If DoValidationChecks2() = False Then Exit Sub
        'データ転送
        If UploadData2() = False Then Exit Sub
    End If
    
    Call MsgBox("終了しました。", vbInformation + vbOKOnly, "")

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
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdFileClear_Click()

    List1.Clear
    List2.Clear

End Sub

Private Sub cmdSelectFolder_Click()

    Dim strFolder As String

    On Error GoTo cmdSelectFolder_Click_Err

    strFolder = OpenSelectFolderDialog(Me.hwnd)
    If Trim(strFolder) <> "" Then
        File1.Path = strFolder
        File1.Refresh
        Dim i As Long
        For i = 0 To (File1.ListCount - 1)
            List1.AddItem File1.List(i)
            List2.AddItem File1.Path
        Next i
    End If
    
    Exit Sub

cmdSelectFolder_Click_Err:

    Call MsgBox("フォルダ変更クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdSelectFolder_Click_Err")

End Sub

Private Sub File1_Click()

    On Error GoTo File1_Click_Err

    Screen.MousePointer = vbHourglass
        
    Picture1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
    Image1.Picture = Picture1.Picture
    
    Screen.MousePointer = vbDefault

    Exit Sub

File1_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("ファイルリストクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "File1_Click_Err")

End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／１１／２８
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
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "受付データと写真をホームページへ転送"
    
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
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set g_clsAdoSQL = Nothing
    Set g_clsAdoAccess = Nothing
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
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    Dim strSQL As String
    Dim adoDT011 As New ADODB.Recordset

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        txtUYear(0).Text = ""
        txtUMonth(0).Text = ""
        txtUDay(0).Text = ""
        txtUYear(1).Text = ""
        txtUMonth(1).Text = ""
        txtUDay(1).Text = ""
        
        '受付明細データオープン(累積)
        strSQL = "SELECT TOP 1 Odate FROM vw_DT011 ORDER BY Odate DESC"
        adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoDT011.EOF = False Then
            txtUYear(0).Text = Format(adoDT011("Odate"), "yyyy")
            txtUMonth(0).Text = Format(adoDT011("Odate"), "m")
            txtUDay(0).Text = Format(adoDT011("Odate"), "d")
            txtUYear(1).Text = Format(adoDT011("Odate"), "yyyy")
            txtUMonth(1).Text = Format(adoDT011("Odate"), "m")
            txtUDay(1).Text = Format(adoDT011("Odate"), "d")
        End If
        adoDT011.Close
        
        txtYear.Text = ""
        txtMonth.Text = ""
        txtDay.Text = ""
        txtSubtitle.Text = ""
        txtTYear.Text = ""
        txtTMonth.Text = ""
        txtTDay.Text = ""
        txtTHour.Text = ""
        txtTmin.Text = ""
        
        Dim strDir As String
        strDir = ""
        If Dir(strDir, vbDirectory) <> "" Then
            File1.Path = ""
            File1.Refresh
        End If
        
        List1.Clear
        List2.Clear
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
        
    If TabStrip1.Tabs(1).Selected = True Then
        txtUYear(0).SetFocus
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        txtYear.SetFocus
    End If

End Sub

Private Sub List1_Click()

    Dim strFileName As String

    On Error GoTo File1_Click_Err

    Screen.MousePointer = vbHourglass
        
    strFileName = List2.List(List1.ListIndex) & "\" & List1.List(List1.ListIndex)
    If Dir(strFileName) <> "" Then
        Picture1.Picture = LoadPicture(strFileName)
        Image1.Picture = Picture1.Picture
    End If
    
    Screen.MousePointer = vbDefault

    Exit Sub

File1_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("ファイルリストクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "File1_Click_Err")

End Sub

Private Sub TabStrip1_Click()

    On Error Resume Next

    If TabStrip1.Tabs(1).Selected = True Then
        fraUketuke.Visible = True
        fraJpg.Visible = False
        txtUYear(0).SetFocus
        DoEvents
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        fraUketuke.Visible = False
        fraJpg.Visible = True
        txtYear.SetFocus
        DoEvents
    End If

End Sub

Private Sub txtDay_GotFocus()
    
    txtDay.BackColor = FOCUS_STOP_COLOR
  
End Sub

Private Sub txtDay_LostFocus()
    
    txtDay.BackColor = FOCUS_NO_COLOR
  
End Sub

Private Sub txtMonth_GotFocus()
    
    txtMonth.BackColor = FOCUS_STOP_COLOR
  
End Sub

Private Sub txtMonth_LostFocus()
    
    txtMonth.BackColor = FOCUS_NO_COLOR
  
End Sub

Private Sub txtSubtitle_GotFocus()
    
    txtSubtitle.BackColor = FOCUS_STOP_COLOR
  
End Sub

Private Sub txtSubtitle_LostFocus()
    
    txtSubtitle.BackColor = FOCUS_NO_COLOR
  
End Sub

Private Sub txtTDay_GotFocus()
    
    txtTDay.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub txtTDay_LostFocus()
    
    txtTDay.BackColor = FOCUS_NO_COLOR
 
End Sub

Private Sub txtTHour_GotFocus()
    
    txtTHour.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtTHour_LostFocus()
    
    txtTHour.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtTmin_GotFocus()
    
    txtTmin.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtTmin_LostFocus()
    
    txtTmin.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtTMonth_GotFocus()
    
    txtTMonth.BackColor = FOCUS_STOP_COLOR
 
End Sub

Private Sub txtTMonth_LostFocus()
    
    txtTMonth.BackColor = FOCUS_NO_COLOR
 
End Sub

Private Sub txtTYear_GotFocus()
    
    txtTYear.BackColor = FOCUS_STOP_COLOR
  
End Sub

Private Sub txtTYear_LostFocus()
    
    txtTYear.BackColor = FOCUS_NO_COLOR
 
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

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Function DoValidationChecks1() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks1_Err

    If Trim(txtUYear(0).Text) = "" Then
        txtUYear(0).SetFocus
        strErrMsg = "対象年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtUMonth(0).Text) = "" Then
        txtUMonth(0).SetFocus
        strErrMsg = "対象年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtUDay(0).Text) = "" Then
        txtUDay(0).SetFocus
        strErrMsg = "対象年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtUYear(1).Text) = "" Then
        txtUYear(1).SetFocus
        strErrMsg = "対象年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtUMonth(1).Text) = "" Then
        txtUMonth(1).SetFocus
        strErrMsg = "対象年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtUDay(1).Text) = "" Then
        txtUDay(1).SetFocus
        strErrMsg = "対象年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks1 = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks1 = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks1_Err:

    DoValidationChecks1 = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks1_Err")

End Function

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Function DoValidationChecks2() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks2_Err

    If Trim(txtYear.Text) = "" Then
        txtYear.SetFocus
        strErrMsg = "開催日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtMonth.Text) = "" Then
        txtMonth.SetFocus
        strErrMsg = "開催日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtDay.Text) = "" Then
        txtDay.SetFocus
        strErrMsg = "開催日を入力してください。"
        GoTo ErrorTrap:
    End If
    If IsDate(txtYear.Text & "/" & txtMonth.Text & "/" & txtDay.Text) = False Then
        txtDay.SetFocus
        strErrMsg = "正しい開催日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtTYear.Text) = "" Then
        txtTYear.SetFocus
        strErrMsg = "撮影日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtTMonth.Text) = "" Then
        txtTMonth.SetFocus
        strErrMsg = "撮影日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtTDay.Text) = "" Then
        txtTDay.SetFocus
        strErrMsg = "撮影日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtTHour.Text) = "" Then
        txtTHour.SetFocus
        strErrMsg = "撮影日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtTmin.Text) = "" Then
        txtTmin.SetFocus
        strErrMsg = "撮影日を入力してください。"
        GoTo ErrorTrap:
    End If
    If IsDate(txtTYear.Text & "/" & txtTMonth.Text & "/" & txtTDay.Text & " " & txtTHour.Text & ":" & txtTmin.Text) = False Then
        txtTDay.SetFocus
        strErrMsg = "正しい撮影日を入力してください。"
        GoTo ErrorTrap:
    End If
    If List1.ListCount <= 0 Then
        List1.SetFocus
        strErrMsg = "ファイルが選択されていません。"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks2 = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks2 = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks2_Err:

    DoValidationChecks2 = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks2_Err")

End Function

'目　的　　：ホームページへデータアップロード(受付データ)
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Function UploadData1() As Boolean

    Dim strSQL As String
    Dim adoDT010 As New ADODB.Recordset
    Dim adoDT011 As New ADODB.Recordset
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim strFileNm As String
    Dim strFTPServer As String
    Dim strFTPDir As String
    Dim strCommand As String

    On Error GoTo UploadData1_Err
    
    UploadData1 = False
    
    Screen.MousePointer = vbHourglass
    
    '対象年月日
    strDateFrom = txtUYear(0).Text & "/" & Format(txtUMonth(0).Text, "00") & "/" & Format(txtUDay(0).Text, "00")
    strDateTo = txtUYear(1).Text & "/" & Format(txtUMonth(1).Text, "00") & "/" & Format(txtUDay(1).Text, "00")
    
    '受付データオープン(累積含む)
    strSQL = "SELECT Odate FROM vw_DT011 WHERE Odate BETWEEN '" & strDateFrom & "' AND '" & strDateTo & "'" & " GROUP BY Odate ORDER BY Odate"
    adoDT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT010.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT010.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If

'********** FTP設定 **********
    
    With Inet1
'        .UserName = "yawas000"
'        .Password = "LceJzJQ6"
        .UserName = "yawaseueki"
        .Password = "36_Yawase_2335"
        .RequestTimeout = 60
        .Protocol = icFTP
'        strFTPServer = "ftp://yawas000:LceJzJQ6@210.172.178.59"
'        strFTPServer = "ftp://yawaseueki:yawaseueki01@120.88.61.216"
'        strFTPServer = "ftp://yawaseueki:yawaseueki01@60.43.207.182"
'        strFTPServer = "ftp://yawaseueki:36_Yawase_2335@60.43.207.182"
        strFTPServer = "ftp://yawaseueki:36_Yawase_2335@www.yawaseueki.co.jp"
'        strFTPDir = "/web/info/data/"
        strFTPDir = "/www/htdocs/info/data/"
    End With

    Do While Not adoDT010.EOF
        strFileNm = "c:\" & Format$(adoDT010("Odate"), "yyyymmdd") & ".txt"
        Open strFileNm For Output As #1
        
        '受付明細データオープン(累積含む)
        strSQL = "SELECT Iname,COALESCE(SUM(Qty),0) AS Qty_Total FROM vw_DT011"
        strSQL = strSQL & " WHERE Odate = '" & adoDT010("Odate") & "'"
        strSQL = strSQL & " GROUP BY Iname ORDER BY Iname"
        adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT011.EOF
            Print #1, """" & Trim(adoDT011("Iname")) & """," & adoDT011("Qty_Total")
                
            adoDT011.MoveNext
        Loop
        adoDT011.Close
        
        Close #1
        
        'FTPファイル転送
        strCommand = "PUT " & strFileNm & " " & strFTPDir & Format$(adoDT010("Odate"), "yyyymmdd") & ".txt"
        Inet1.Execute strFTPServer, strCommand
        'ビジーの間は待つ
        Do While Inet1.StillExecuting = True
            DoEvents
            If frmCount.g_blnCancel Then GoTo UploadData1_Cancel:
        Loop
        
        On Error Resume Next
        Kill strFileNm
        On Error GoTo UploadData1_Err
        
        adoDT010.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo UploadData1_Cancel:
    Loop
    adoDT010.Close
    
'********** FTP切断 **********
    
    strCommand = "CLOSE"
    Inet1.Execute , strCommand
    
    UploadData1 = True
    
UploadData1_Exit:
    Close
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    DoEvents
    
    Exit Function

UploadData1_Cancel:

    GoTo UploadData1_Exit:

UploadData1_Err:

    UploadData1 = False
    Call MsgBox("ホームページへデータアップロードエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "UploadData1_Err")
    GoTo UploadData1_Exit:

End Function

'目　的　　：ホームページへデータアップロード(写真)
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００６／０１／１３
'更新履歴　：
'
Private Function UploadData2() As Boolean

    Dim intIndex1 As Integer
    Dim strOdate As String
    Dim strTakeDate As String
    Dim strFTPServer As String
    Dim strFTPDir As String
    Dim strCommand As String
    Dim strFTPDirTmp() As String
    
    Dim vntGetData As Variant
    Dim strSplit() As String
    Dim lngBufSize As Long
    Dim lngCounter As Long

    On Error GoTo UploadData2_Err
    
    UploadData2 = False
    
    Screen.MousePointer = vbHourglass
    
    '開催日
    strOdate = txtYear.Text & "/" & Format(txtMonth.Text, "00") & "/" & Format(txtDay.Text, "00")
    '撮影日
    strTakeDate = txtTYear.Text & "/" & Format(txtTMonth.Text, "00") & "/" & Format(txtTDay.Text, "00") & " " & txtTHour.Text & ":" & txtTmin.Text
    
    With frmCount
        .fpProgressBar1.Value = 0
        .fpProgressBar1.Max = List1.ListCount
        .Show
        Me.Enabled = False
        DoEvents
    End With
    
'********** FTP設定 **********
    
    With Inet1
'        .UserName = "yawas000"
'        .Password = "LceJzJQ6"
        .UserName = "yawaseueki"
        .Password = "36_Yawase_2335"
        .RequestTimeout = 60
        .Protocol = icFTP
'        strFTPServer = "ftp://yawas000:LceJzJQ6@210.172.178.59"
'        strFTPServer = "ftp://yawaseueki:yawaseueki01@120.88.61.216"
'        strFTPServer = "ftp://yawaseueki:36_Yawase_2335@60.43.207.182"
        strFTPServer = "ftp://yawaseueki:36_Yawase_2335@www.yawaseueki.co.jp"
        
'        strFTPDir = "/web/info/jpg/"
        strFTPDir = "/www/htdocs/info/jpg/"
    End With
    
'********** フォルダ内の全ファイル削除 **********
    
    'フォルダ一覧表示
    strCommand = "DIR " & strFTPDir
    Inet1.Execute strFTPServer, strCommand
    'ビジーの間は待つ
    Do While Inet1.StillExecuting = True
        DoEvents
        If frmCount.g_blnCancel Then GoTo UploadData2_Cancel:
    Loop
    vntGetData = Inet1.GetChunk(1024, icString)
    '戻り値が長さ0の文字列だったらループを抜ける
    If Len(vntGetData) <> 0 Then
        '取得したデータを元に配列を作成
        strSplit = Split(vntGetData, vbNewLine)
        '配列の各要素を調べる
        For lngCounter = 0 To UBound(strSplit)
            '長さ0の配列要素を除く
            If Len(strSplit(lngCounter)) <> 0 Then
                'フォルダ以外
                If right(strSplit(lngCounter), 1) <> "/" Then
'                    Debug.Print strSplit(lngCounter)
                    '削除
                    strCommand = "DELETE " & strFTPDir & strSplit(lngCounter)
                    Inet1.Execute strFTPServer, strCommand
                    'ビジーの間は待つ
                    Do While Inet1.StillExecuting = True
                        DoEvents
                        If frmCount.g_blnCancel Then GoTo UploadData2_Cancel:
                    Loop
                End If
            End If
        Next
    End If

'********** FTPファイル転送と画像リスト作成 **********
    
    Open "c:\list.txt" For Output As #1
    
    For intIndex1 = 0 To (List1.ListCount - 1)
        FileCopy List2.List(intIndex1) & "\" & List1.List(intIndex1), "C:\" & List1.List(intIndex1)
    
        'FTPファイル転送
        strCommand = "PUT " & "c:\" & List1.List(intIndex1) & " " & strFTPDir & List1.List(intIndex1)
        Inet1.Execute strFTPServer, strCommand
        'ビジーの間は待つ
        Do While Inet1.StillExecuting = True
            DoEvents
            If frmCount.g_blnCancel Then GoTo UploadData2_Cancel:
        Loop
        
        On Error Resume Next
        Kill "C:\" & List1.List(intIndex1)
        On Error GoTo UploadData2_Err
        
        Print #1, List1.List(intIndex1)

        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo UploadData2_Cancel:
    Next intIndex1

    Close #1
    
    '画像リストFTPファイル転送
    'strCommand = "PUT c:\" & "list.txt" & " /web/info/list.txt"
    strCommand = "PUT c:\" & "list.txt" & " /www/htdocs/info/list.txt"
    Inet1.Execute strFTPServer, strCommand
    'ビジーの間は待つ
    Do While Inet1.StillExecuting = True
        DoEvents
        If frmCount.g_blnCancel Then GoTo UploadData2_Cancel:
    Loop

'********** コメント欄(開催日など)の作成 **********
    
    Open "c:\comment.txt" For Output As #1
    Print #1, Format$(strOdate, "m月d日")
    Print #1, txtSubtitle.Text
    Print #1, Format$(strTakeDate, "m月d日 hh:nn")
    Close #1
    
    'FTPファイル転送
    'strCommand = "PUT c:\" & "comment.txt" & " /web/info/comment.txt"
    strCommand = "PUT c:\" & "comment.txt" & " /www/htdocs/info/comment.txt"
    Inet1.Execute strFTPServer, strCommand
    'ビジーの間は待つ
    Do While Inet1.StillExecuting = True
        DoEvents
        If frmCount.g_blnCancel Then GoTo UploadData2_Cancel:
    Loop
    
'********** FTP切断 **********
    
    strCommand = "CLOSE"
    Inet1.Execute , strCommand
    
    UploadData2 = True
    
UploadData2_Exit:
    Close
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    DoEvents
    
    Exit Function

UploadData2_Cancel:

    GoTo UploadData2_Exit:

UploadData2_Err:

    UploadData2 = False
    Call MsgBox("ホームページへデータアップロードエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "UploadData2_Err")
    GoTo UploadData2_Exit:

End Function

Private Sub txtYear_GotFocus()
    
    txtYear.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtYear_LostFocus()
    
    txtYear.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：ホームページへデータアップロード(受付データ)
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／１１／２８
'更新履歴　：
'
Private Function UploadData1_old() As Boolean

    Dim strSQL As String
    Dim adoDT011 As New ADODB.Recordset
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim MySqlCn As Variant

    On Error GoTo UploadData1_old_Err
    
    UploadData1_old = False
    
    Screen.MousePointer = vbHourglass
    
    '対象年月日
    strDateFrom = txtUYear(0).Text & "/" & Format(txtUMonth(0).Text, "00") & "/" & Format(txtUDay(0).Text, "00")
    strDateTo = txtUYear(1).Text & "/" & Format(txtUMonth(1).Text, "00") & "/" & Format(txtUDay(1).Text, "00")
    
    '受付明細データオープン(累積)
    strSQL = "SELECT Odate,Iname,SUM(Qty) AS Qty_Total FROM vw_DT011 WHERE Odate BETWEEN '" & strDateFrom & "' AND '" & strDateTo & "'" & " GROUP BY Odate,Iname ORDER BY Odate,Iname"
    adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT011.EOF = False Then
        With frmCount
            .fpProgressBar1.Value = 0
            .fpProgressBar1.Max = adoDT011.RecordCount
            .Show
            Me.Enabled = False
        End With
    End If
    
    'MySql ODBC接続
    Set MySqlCn = CreateObject("ADODB.Connection")
'    MySqlCn.Open "dsn=YAWASEHP;uid=root;pwd="
    MySqlCn.Open "dsn=YAWASEHP;"
'    MySqlCn.CursorLocation = 3

    'トランザクション開始
'    MySqlCn.BeginTrans
    
    'データ削除
    strSQL = "delete from arrival where Odate >= '" & Format(strDateFrom, "yyyy-mm-dd") & "' AND Odate <= '" & Format(strDateTo, "yyyy-mm-dd") & "'"
    MySqlCn.Execute strSQL

    Do While Not adoDT011.EOF
        'データ作成
        strSQL = "insert into arrival(Odate,Iname,Qty)"
        strSQL = strSQL & " values('" & Format$(adoDT011("Odate"), "yyyy-mm-dd") & "',"
        strSQL = strSQL & " '" & adoDT011("Iname") & "',"
        strSQL = strSQL & " " & adoDT011("Qty_Total") & ")"
        MySqlCn.Execute strSQL
                
        adoDT011.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo UploadData1_old_Cancel:
    Loop
    adoDT011.Close
    
'    MySqlCn.CommitTrans
    MySqlCn.Close
    Set MySqlCn = Nothing
    
    UploadData1_old = True
    
UploadData1_old_Exit:
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    DoEvents
    
    Exit Function

UploadData1_old_Cancel:

    GoTo UploadData1_old_Exit:

UploadData1_old_Err:

    UploadData1_old = False
    Call MsgBox("ホームページへデータアップロードエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "UploadData1_old_Err")
    GoTo UploadData1_old_Exit:

End Function

