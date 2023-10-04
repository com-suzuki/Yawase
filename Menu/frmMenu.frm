VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   6525
   ClientLeft      =   150
   ClientTop       =   330
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   14.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9870
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   5055
      Left            =   3840
      TabIndex        =   24
      Top             =   720
      Width           =   5895
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   420
         Index           =   1
         Left            =   60
         TabIndex        =   25
         Top             =   180
         Width           =   5775
         _Version        =   262145
         _ExtentX        =   10186
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "処　　理　　名"
         ForeColor       =   0
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         TopMargin       =   0
         LabelTop        =   0
         LabelWidth      =   99
         LabelHeight     =   27
         LabelLeft       =   143
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
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1140
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":0CFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":0D68
         Key             =   "frmMenu.frx":0D86
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":0DBA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":0E28
         Key             =   "frmMenu.frx":0E46
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1980
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":0E7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":0EE8
         Key             =   "frmMenu.frx":0F06
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":0F3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":0FA8
         Key             =   "frmMenu.frx":0FC6
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2820
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":0FFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1068
         Key             =   "frmMenu.frx":1086
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":10BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1128
         Key             =   "frmMenu.frx":1146
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   3660
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":117A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":11E8
         Key             =   "frmMenu.frx":1206
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":123A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":12A8
         Key             =   "frmMenu.frx":12C6
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   4500
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":12FA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1368
         Key             =   "frmMenu.frx":1386
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtSubMenu 
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":13BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1428
         Key             =   "frmMenu.frx":1446
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '実線
         Height          =   4335
         Left            =   60
         TabIndex        =   28
         Top             =   660
         Width           =   5775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   5055
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   3675
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   435
         Index           =   0
         Left            =   60
         TabIndex        =   23
         Top             =   180
         Width           =   3555
         _Version        =   262145
         _ExtentX        =   6271
         _ExtentY        =   767
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         TopMargin       =   0
         LabelTop        =   0
         LabelWidth      =   76
         LabelHeight     =   27
         LabelLeft       =   80
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
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":147A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":14E8
         Key             =   "frmMenu.frx":1506
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   1140
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":154A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":15B8
         Key             =   "frmMenu.frx":15D6
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":161A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1688
         Key             =   "frmMenu.frx":16A6
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1980
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":16EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1758
         Key             =   "frmMenu.frx":1776
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":17BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1828
         Key             =   "frmMenu.frx":1846
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   2820
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":188A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":18F8
         Key             =   "frmMenu.frx":1916
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":195A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":19C8
         Key             =   "frmMenu.frx":19E6
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   3660
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":1A2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1A98
         Key             =   "frmMenu.frx":1AB6
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   4080
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":1AFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1B68
         Key             =   "frmMenu.frx":1B86
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imtMenu 
         Height          =   435
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   4500
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         Caption         =   "frmMenu.frx":1BCA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMenu.frx":1C38
         Key             =   "frmMenu.frx":1C56
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   0
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '実線
         Height          =   4335
         Left            =   60
         TabIndex        =   27
         Top             =   660
         Width           =   3555
      End
   End
   Begin CSCaptLib.CSCaption lblTitle 
      Height          =   525
      Left            =   60
      TabIndex        =   21
      Top             =   60
      Width           =   9735
      _Version        =   262145
      _ExtentX        =   17171
      _ExtentY        =   926
      _StockProps     =   79
      ForeColor       =   0
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      TopMargin       =   0
      LabelTop        =   0
      LabelWidth      =   170
      LabelHeight     =   35
      LabelLeft       =   239
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
      Height          =   5235
      Index           =   2
      Left            =   60
      TabIndex        =   26
      Top             =   660
      Width           =   9735
      _Version        =   262145
      _ExtentX        =   17171
      _ExtentY        =   9234
      _StockProps     =   79
      ForeColor       =   0
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
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
      LabelHeight     =   32
      LabelLeft       =   283
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
   Begin CSCmdLibCtl.CSCmdBtn cmdExit 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   7200
      TabIndex        =   20
      Top             =   6000
      Width           =   2595
      _Version        =   262145
      _ExtentX        =   4577
      _ExtentY        =   767
      _StockProps     =   15
      Caption         =   "システム終了"
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      SizePicture     =   -1  'True
      OwnPicDisabled  =   0   'False
      CaptionPosition =   3
      rPic.left       =   27
      rPic.top        =   4
      rPic.right      =   21
      rPic.bottom     =   21
      rText.left      =   71
      rText.top       =   5
      rText.right     =   171
      rText.bottom    =   25
      Picture         =   "frmMenu.frx":1C9A
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_ROW = 10                          'メニュー表示行数
Const SELCT_COLOR1 = &HFFFFFF               '背景色(白)
Const SELCT_COLOR2 = &H80000012             '背景色(黒)
Const MENU_DAT_FILENAME = "Menu.txt"

Private m_intLastMenuIndex As Integer
Private m_intLastSubMenuIndex As Integer

Private Type typSubMenu
    MenuIndex As Integer
    SubMenuIndex As Integer
    SubMenuName As String
    SubMenuProgramCmd As String
    SubMenuProgramName As String
End Type
Private m_typSubMenu() As typSubMenu

'現在アクティブなウィンドウを取得するAPI
Private Declare Function GetForegroundWindow Lib "USER32" () As Long

Private Function ShellProgram(strPgName As String, strTitleName As String) As Boolean

    Dim strCaption As String
    
    On Error GoTo ShellProgram_Err
    
    Screen.MousePointer = vbHourglass
    
    'プログラムの実行
    strCaption = SYSTEM_NAME & "-" & strTitleName
    If Global_IsLoad(vbNullString, strCaption, "", "") = False Then Exit Function
    Call Global_IsLoad(vbNullString, strCaption, g_clsReg.Bin & "\", strPgName)

    Screen.MousePointer = vbDefault

    Exit Function

ShellProgram_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("プログラムの実行エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ShellProgram_Err")

End Function

Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub Form_Load()

    Dim intIndex1 As Integer

    On Error GoTo Form_Load_Err

'    Me.Caption = SYSTEM_NAME & " Ver" & PROGRAM_VERSION
    Me.Caption = "メニュー"
    lblTitle.Caption = SYSTEM_NAME

    '初期化
    m_intLastMenuIndex = 9999
    m_intLastSubMenuIndex = 9999

    'フォームクリア
    For intIndex1 = 0 To MAX_ROW - 1
        imtMenu(intIndex1).Text = ""
        imtSubMenu(intIndex1).Text = ""
    Next intIndex1

    If ReadMenuItem() = False Then End

    Exit Sub
    
Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")

End Sub

Private Function ReadMenuItem() As Boolean

    Dim intFreefile1 As Integer
    Dim strFileName As String
    Dim strBuff1 As String
    Dim intFildCnt As Integer
    Dim varFildData() As Variant
    Dim intMenuIndex As Integer

    On Error GoTo ReadMenuItem_Err

    ReadMenuItem = False

    'ファイル名
    strFileName = App.Path & "\" & MENU_DAT_FILENAME
        
    'ファイルが存在するかチェック
    If Dir(strFileName) = "" Then
        Call MsgBox("メニュー設定データがありません。", vbOKOnly + vbInformation, "")
        ReadMenuItem = True
        Exit Function
    End If
    
    '初期化
    intMenuIndex = 1
    Erase m_typSubMenu
    
    intFreefile1 = FreeFile
    Open strFileName For Input Shared As intFreefile1
    Do While Not EOF(intFreefile1)
        Line Input #intFreefile1, strBuff1
        'ＣＳＶ取り込み変換
        Call S_CSVtoTEXT(strBuff1, varFildData, intFildCnt)
        Select Case LCase(varFildData(1))
            Case "a":   'メニュー名表示
                imtMenu(varFildData(2) - 1).Text = Trim(varFildData(3))
            Case "b":   '処理名の記憶
                ReDim Preserve m_typSubMenu(intMenuIndex)
                
                m_typSubMenu(intMenuIndex).MenuIndex = varFildData(2)
                m_typSubMenu(intMenuIndex).SubMenuIndex = varFildData(3)
                m_typSubMenu(intMenuIndex).SubMenuName = varFildData(4)
                m_typSubMenu(intMenuIndex).SubMenuProgramCmd = varFildData(5)
                m_typSubMenu(intMenuIndex).SubMenuProgramName = varFildData(6)
                
                intMenuIndex = intMenuIndex + 1
            Case Else
        End Select
    Loop
    Close intFreefile1
    
    ReadMenuItem = True
    
    Exit Function

ReadMenuItem_Err:

    Close
    ReadMenuItem = False
    Call MsgBox("メニュー読み込みエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ReadMenuItem_Err")

End Function

Private Function DispSubMenuItem(intMenu As Integer) As Boolean

    Dim intIndex1 As Integer

    On Error GoTo DispSubMenuItem_Err

    DispSubMenuItem = False

    'サブメニューのクリア
    For intIndex1 = 0 To MAX_ROW - 1
        imtSubMenu(intIndex1).Text = ""
        imtSubMenu(intIndex1).BackColor = SELCT_COLOR1
        imtSubMenu(intIndex1).ForeColor = SELCT_COLOR2
    Next intIndex1

    For intIndex1 = 1 To UBound(m_typSubMenu)
        If m_typSubMenu(intIndex1).MenuIndex = intMenu Then
            imtSubMenu(m_typSubMenu(intIndex1).SubMenuIndex - 1).Text = m_typSubMenu(intIndex1).SubMenuName
        End If
    Next intIndex1
                
    DispSubMenuItem = True
    
    Exit Function

DispSubMenuItem_Err:

    DispSubMenuItem = False
    Call MsgBox("サブメニュー表示エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DispSubMenuItem_Err")

End Function

Private Function SelectMenuItem(intSubMenu As Integer) As Boolean

    Dim intIndex1 As Integer
    Dim intMenu As Integer

    On Error GoTo SelectMenuItem_Err

    SelectMenuItem = False

    '選択されているメニューを探す
    intMenu = m_intLastMenuIndex + 1

    For intIndex1 = 1 To UBound(m_typSubMenu)
        If m_typSubMenu(intIndex1).MenuIndex = intMenu And m_typSubMenu(intIndex1).SubMenuIndex = intSubMenu Then
            Call ShellProgram(m_typSubMenu(intIndex1).SubMenuProgramCmd, m_typSubMenu(intIndex1).SubMenuProgramName)
            Exit For
        End If
    Next intIndex1
                
    SelectMenuItem = True
    
    Exit Function

SelectMenuItem_Err:

    SelectMenuItem = False
    Call MsgBox("メニュー選択エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "SelectMenuItem_Err")

End Function

Private Sub imtMenu_GotFocus(Index As Integer)

    Dim intIndex1 As Integer

    On Error GoTo imtMenu_GotFocus_Err

    '背景色クリア
    For intIndex1 = 0 To (MAX_ROW - 1)
        imtMenu(intIndex1).BackColor = SELCT_COLOR1
        imtMenu(intIndex1).ForeColor = SELCT_COLOR2
    Next intIndex1

    imtMenu(Index).SelStart = 0
    imtMenu(Index).BackColor = SELCT_COLOR2
    imtMenu(Index).ForeColor = SELCT_COLOR1
    m_intLastMenuIndex = Index
    
    Call DispSubMenuItem(Index + 1)

    Exit Sub

imtMenu_GotFocus_Err:
    
    Call MsgBox("メニューフォーカス取得時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtMenu_GotFocus_Err")

End Sub

Private Sub imtMenu_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    On Error GoTo imtSubMenu_KeyDown_Err

    If KeyCode = vbKeyUp Then
        If Index > 0 Then
            imtMenu(Index - 1).SetFocus
        End If
    End If
    If KeyCode = vbKeyDown Then
        If (Index + 1) < MAX_ROW Then
            If Trim(imtMenu(Index + 1).Text) <> "" Then
                imtMenu(Index + 1).SetFocus
            End If
        End If
    End If
    If KeyCode = vbKeyRight Then
        imtSubMenu(0).SetFocus
    End If

    Exit Sub

imtSubMenu_KeyDown_Err:

    Call MsgBox("メニューキーダウンエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtSubMenu_KeyDown_Err")

End Sub

Private Sub imtMenu_LostFocus(Index As Integer)

    imtMenu(Index).BackColor = SELCT_COLOR1
    imtMenu(Index).ForeColor = SELCT_COLOR2
    
End Sub

Private Sub imtMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo imtMenu_MouseMove_Err
        
    '現在アクティブなウインドウが、自フォームでない時は処理無し
    If GetForegroundWindow() <> Me.hwnd Then Exit Sub
    
    If Trim(imtMenu(Index).Text) = "" Then Exit Sub
    imtMenu(Index).SetFocus
    DoEvents

    Exit Sub

imtMenu_MouseMove_Err:

    Call MsgBox("メニューマウス移動時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtMenu_MouseMove_Err")

End Sub

Private Sub imtSubMenu_Click(Index As Integer)

    Call SelectMenuItem(Index + 1)

End Sub

Private Sub imtSubMenu_DblClick(Index As Integer)

    Call SelectMenuItem(Index + 1)
    
End Sub

Private Sub imtSubMenu_GotFocus(Index As Integer)
    
    On Error GoTo imtSubMenu_GotFocus_Err

    'メニューの背景色を変える
    If m_intLastMenuIndex < MAX_ROW And m_intLastMenuIndex >= 0 Then
        imtMenu(m_intLastMenuIndex).BackColor = SELCT_COLOR2
        imtMenu(m_intLastMenuIndex).ForeColor = SELCT_COLOR1
    End If

    imtSubMenu(Index).SelStart = 0
    imtSubMenu(Index).BackColor = SELCT_COLOR2
    imtSubMenu(Index).ForeColor = SELCT_COLOR1
    
    Exit Sub
    
imtSubMenu_GotFocus_Err:
    
    Call MsgBox("サブメニューフォーカス取得時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtSubMenu_GotFocus_Err")

End Sub

Private Sub imtSubMenu_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    On Error GoTo imtSubMenu_KeyDown_Err

    If KeyCode = vbKeyUp Then
        If Index > 0 Then
            imtSubMenu(Index - 1).SetFocus
        End If
    End If
    If KeyCode = vbKeyDown Then
        If (Index + 1) < MAX_ROW Then
            If Trim(imtSubMenu(Index + 1).Text) <> "" Then
                imtSubMenu(Index + 1).SetFocus
            End If
        End If
    End If
    If KeyCode = vbKeyLeft Then
        If m_intLastMenuIndex < MAX_ROW And m_intLastMenuIndex >= 0 Then
            imtMenu(m_intLastMenuIndex).SetFocus
        Else
            m_intLastMenuIndex = 0
            imtMenu(m_intLastMenuIndex).SetFocus
        End If
    End If
    If KeyCode = vbKeyReturn Then
        Call SelectMenuItem(Index + 1)
    End If
    
    Exit Sub

imtSubMenu_KeyDown_Err:

    Call MsgBox("メニューキーダウンエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtSubMenu_KeyDown_Err")

End Sub

Private Sub imtSubMenu_LostFocus(Index As Integer)

    imtSubMenu(Index).BackColor = SELCT_COLOR1
    imtSubMenu(Index).ForeColor = SELCT_COLOR2
    
End Sub

Private Sub imtSubMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    On Error GoTo imtSubMenu_MouseMove_Err
        
    '現在アクティブなウインドウが、自フォームでない時は処理無し
    If GetForegroundWindow() <> Me.hwnd Then Exit Sub
        
    If Trim(imtSubMenu(Index).Text) = "" Then Exit Sub
    If m_intLastSubMenuIndex = Index Then Exit Sub
    imtSubMenu(Index).SetFocus
    DoEvents

    Exit Sub

imtSubMenu_MouseMove_Err:

    Call MsgBox("サブメニューマウス移動時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtSubMenu_MouseMove_Err")

End Sub
