VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmMt010 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   ClientHeight    =   9975
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMt010.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   13500
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame fra 
      Height          =   735
      Left            =   60
      TabIndex        =   38
      Top             =   9180
      Width           =   13335
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         Height          =   495
         Left            =   60
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "âÊñ ∏ÿ±(F8)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Picture         =   "frmMt010.frx":0CFA
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   11520
         TabIndex        =   34
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "èIóπ(F9)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Picture         =   "frmMt010.frx":0D16
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   9780
         TabIndex        =   33
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "é¿çs(F12)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Picture         =   "frmMt010.frx":0E70
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   13740
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmMt010.frx":12C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt010.frx":1330
      Key             =   "frmMt010.frx":134E
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
      Left            =   13980
      TabIndex        =   35
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmMt010.frx":1392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt010.frx":1400
      Key             =   "frmMt010.frx":141E
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
   Begin VB.Frame fraMeisai 
      Height          =   9135
      Left            =   60
      TabIndex        =   36
      Top             =   60
      Width           =   13335
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'Ç»Çµ
         Height          =   375
         Left            =   8940
         TabIndex        =   47
         Top             =   2160
         Width           =   4335
         Begin VB.OptionButton optBfraction 
            Caption         =   "êÿÇËè„Ç∞"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2880
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   13
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optBfraction 
            Caption         =   "êÿÇËéÃÇƒ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   0
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   11
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optBfraction 
            Caption         =   "éléÃå‹ì¸"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1440
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   12
            Top             =   0
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'Ç»Çµ
         Height          =   495
         Left            =   2220
         TabIndex        =   46
         Top             =   1620
         Width           =   4335
         Begin VB.OptionButton optSfraction 
            Caption         =   "éléÃå‹ì¸"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1440
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   5
            Top             =   60
            Width           =   1395
         End
         Begin VB.OptionButton optSfraction 
            Caption         =   "êÿÇËéÃÇƒ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   0
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   4
            Top             =   60
            Width           =   1395
         End
         Begin VB.OptionButton optSfraction 
            Caption         =   "êÿÇËè„Ç∞"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2880
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   6
            Top             =   60
            Width           =   1395
         End
      End
      Begin imText6Ctl.imText txtCompany 
         Height          =   405
         Left            =   2220
         TabIndex        =   1
         Top             =   180
         Width           =   7515
         _Version        =   65536
         _ExtentX        =   13256
         _ExtentY        =   714
         Caption         =   "frmMt010.frx":1462
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":14D0
         Key             =   "frmMt010.frx":14EE
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
         Index           =   3
         Left            =   60
         TabIndex        =   37
         Top             =   180
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "é©é–ñº"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   48
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
         TabIndex        =   39
         Top             =   660
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "èoïié“à€éùä«óùîÔ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelWidth      =   120
         LabelHeight     =   25
         LabelLeft       =   10
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
      Begin imNumber6Ctl.imNumber imnSkeep 
         Height          =   375
         Left            =   2220
         TabIndex        =   2
         Top             =   660
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":1532
         Caption         =   "frmMt010.frx":1552
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":15C0
         Keys            =   "frmMt010.frx":15DE
         Spin            =   "frmMt010.frx":1628
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   6780
         TabIndex        =   40
         Top             =   660
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "îÉéÂà€éùä«óùîÔ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   17
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
      Begin imNumber6Ctl.imNumber imnBkeep 
         Height          =   375
         Left            =   8940
         TabIndex        =   9
         Top             =   660
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":1650
         Caption         =   "frmMt010.frx":1670
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":16DE
         Keys            =   "frmMt010.frx":16FC
         Spin            =   "frmMt010.frx":1746
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   4
         Left            =   60
         TabIndex        =   41
         Top             =   1140
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "èoïié“éËêîóøó¶"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   17
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
      Begin imNumber6Ctl.imNumber imnSrate 
         Height          =   375
         Left            =   2220
         TabIndex        =   3
         Top             =   1140
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":176E
         Caption         =   "frmMt010.frx":178E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":17FC
         Keys            =   "frmMt010.frx":181A
         Spin            =   "frmMt010.frx":1864
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   5
         Left            =   6780
         TabIndex        =   42
         Top             =   1140
         Visible         =   0   'False
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "îÉéÂéËêîóøó¶"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   25
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
      Begin imNumber6Ctl.imNumber imnBrate 
         Height          =   375
         Left            =   8940
         TabIndex        =   10
         Top             =   1140
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":188C
         Caption         =   "frmMt010.frx":18AC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":191A
         Keys            =   "frmMt010.frx":1938
         Spin            =   "frmMt010.frx":1982
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   6
         Left            =   60
         TabIndex        =   43
         Top             =   1680
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "èoïié“í[êîèàóù"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   17
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
         Left            =   6780
         TabIndex        =   44
         Top             =   2160
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "îÉéÂí[êîèàóù"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   25
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
         Left            =   60
         TabIndex        =   45
         Top             =   2640
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "è¨êÿéËéËêîóø"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   25
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
      Begin imNumber6Ctl.imNumber imnCrate 
         Height          =   375
         Left            =   2220
         TabIndex        =   8
         Top             =   2640
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":19AA
         Caption         =   "frmMt010.frx":19CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":1A38
         Keys            =   "frmMt010.frx":1A56
         Spin            =   "frmMt010.frx":1AA0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   53
         Top             =   2160
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "èoïié“í[êîèàóùíPà "
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelWidth      =   135
         LabelHeight     =   25
         LabelLeft       =   2
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
      Begin imNumber6Ctl.imNumber imnSRounding 
         Height          =   375
         Left            =   2220
         TabIndex        =   7
         Top             =   2160
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":1AC8
         Caption         =   "frmMt010.frx":1AE8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":1B56
         Keys            =   "frmMt010.frx":1B74
         Spin            =   "frmMt010.frx":1BBE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   9
         Left            =   6780
         TabIndex        =   55
         Top             =   2640
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "îÉéÂí[êîèàóùíPà "
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelWidth      =   120
         LabelHeight     =   25
         LabelLeft       =   10
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
      Begin imNumber6Ctl.imNumber imnBRounding 
         Height          =   375
         Left            =   8940
         TabIndex        =   14
         Top             =   2640
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":1BE6
         Caption         =   "frmMt010.frx":1C06
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":1C74
         Keys            =   "frmMt010.frx":1C92
         Spin            =   "frmMt010.frx":1CDC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   10
         Left            =   60
         TabIndex        =   57
         Top             =   3360
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ésèÍëçñ êœ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   32
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
      Begin imNumber6Ctl.imNumber imnTm2 
         Height          =   375
         Left            =   2220
         TabIndex        =   15
         Top             =   3360
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":1D04
         Caption         =   "frmMt010.frx":1D24
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":1D92
         Keys            =   "frmMt010.frx":1DB0
         Spin            =   "frmMt010.frx":1DF2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
         MinValue        =   -999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   11
         Left            =   60
         TabIndex        =   59
         Top             =   3840
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ñÏäOîÑèÍñ êœ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   25
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
      Begin imNumber6Ctl.imNumber imnOm2 
         Height          =   375
         Left            =   2220
         TabIndex        =   16
         Top             =   3840
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":1E1A
         Caption         =   "frmMt010.frx":1E3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":1EA8
         Keys            =   "frmMt010.frx":1EC6
         Spin            =   "frmMt010.frx":1F08
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
         MinValue        =   -999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   12
         Left            =   60
         TabIndex        =   61
         Top             =   4320
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "âÆì‡îÑèÍñ êœ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   25
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
      Begin imNumber6Ctl.imNumber imnIm2 
         Height          =   375
         Left            =   2220
         TabIndex        =   17
         Top             =   4320
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":1F30
         Caption         =   "frmMt010.frx":1F50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":1FBE
         Keys            =   "frmMt010.frx":1FDC
         Spin            =   "frmMt010.frx":201E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
         MinValue        =   -999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   13
         Left            =   60
         TabIndex        =   63
         Top             =   4800
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "íìé‘èÍñ êœ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   32
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
      Begin imNumber6Ctl.imNumber imnCm2 
         Height          =   375
         Left            =   2220
         TabIndex        =   18
         Top             =   4800
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":2046
         Caption         =   "frmMt010.frx":2066
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":20D4
         Keys            =   "frmMt010.frx":20F2
         Spin            =   "frmMt010.frx":2134
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
         MinValue        =   -999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   14
         Left            =   60
         TabIndex        =   65
         Top             =   5280
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "éññ±èäñ êœ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   32
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
      Begin imNumber6Ctl.imNumber imnBm2 
         Height          =   375
         Left            =   2220
         TabIndex        =   19
         Top             =   5280
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":215C
         Caption         =   "frmMt010.frx":217C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":21EA
         Keys            =   "frmMt010.frx":2208
         Spin            =   "frmMt010.frx":224A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
         MinValue        =   -999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   15
         Left            =   4080
         TabIndex        =   67
         Top             =   3360
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "îÑóßã‡äzíPà "
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   25
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
      Begin imNumber6Ctl.imNumber imnUp 
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   3360
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":2272
         Caption         =   "frmMt010.frx":2292
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2300
         Keys            =   "frmMt010.frx":231E
         Spin            =   "frmMt010.frx":2360
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
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
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   16
         Left            =   4080
         TabIndex        =   69
         Top             =   3840
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ëOâÒèàóùì˙"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   32
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
         Index           =   17
         Left            =   4080
         TabIndex        =   70
         Top             =   4320
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "îÉéÂì`ï[ÉÅÉÇ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelWidth      =   84
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
      Begin imText6Ctl.imText imtMemo 
         Height          =   825
         Left            =   6240
         TabIndex        =   24
         Top             =   4320
         Width           =   6915
         _Version        =   65536
         _ExtentX        =   12197
         _ExtentY        =   1455
         Caption         =   "frmMt010.frx":2388
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":23F6
         Key             =   "frmMt010.frx":2414
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
      Begin imNumber6Ctl.imNumber imnPdate 
         Height          =   375
         Index           =   0
         Left            =   6240
         TabIndex        =   21
         Top             =   3840
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":2458
         Caption         =   "frmMt010.frx":2478
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":24E6
         Keys            =   "frmMt010.frx":2504
         Spin            =   "frmMt010.frx":2546
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
         ValueVT         =   2012217349
         Value           =   2099
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnPdate 
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   22
         Top             =   3840
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":256E
         Caption         =   "frmMt010.frx":258E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":25FC
         Keys            =   "frmMt010.frx":261A
         Spin            =   "frmMt010.frx":265C
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
         ValueVT         =   2012217349
         Value           =   12
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnPdate 
         Height          =   375
         Index           =   2
         Left            =   8160
         TabIndex        =   23
         Top             =   3840
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":2684
         Caption         =   "frmMt010.frx":26A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2712
         Keys            =   "frmMt010.frx":2730
         Spin            =   "frmMt010.frx":2772
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
         ValueVT         =   5
         Value           =   31
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   18
         Left            =   60
         TabIndex        =   74
         Top             =   6480
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "èZÅ@èä"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   50
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
         Index           =   19
         Left            =   60
         TabIndex        =   75
         Top             =   7380
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ìdòbî‘çÜ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   40
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
         Left            =   2220
         TabIndex        =   26
         Top             =   6480
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   1455
         Caption         =   "frmMt010.frx":279A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2808
         Key             =   "frmMt010.frx":2826
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
         Left            =   2220
         TabIndex        =   27
         Top             =   7380
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   635
         Caption         =   "frmMt010.frx":285A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":28C8
         Key             =   "frmMt010.frx":28E6
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
         Index           =   20
         Left            =   60
         TabIndex        =   76
         Top             =   6000
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "óXï÷î‘çÜ"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   40
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
         Left            =   2220
         TabIndex        =   25
         Top             =   6000
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   635
         Caption         =   "frmMt010.frx":291A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2988
         Key             =   "frmMt010.frx":29A6
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
         Index           =   21
         Left            =   5460
         TabIndex        =   77
         Top             =   7380
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ÇeÇ`Çw"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   55
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
         Left            =   7620
         TabIndex        =   28
         Top             =   7380
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   635
         Caption         =   "frmMt010.frx":29DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2A48
         Key             =   "frmMt010.frx":2A66
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
         Index           =   22
         Left            =   60
         TabIndex        =   78
         Top             =   7800
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ë„ï\é“ñº"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   40
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
         Left            =   2220
         TabIndex        =   29
         Top             =   7800
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frmMt010.frx":2A9A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2B08
         Key             =   "frmMt010.frx":2B26
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
         Index           =   23
         Left            =   60
         TabIndex        =   79
         Top             =   8220
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ÉzÅ[ÉÄÉyÅ[ÉW"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelWidth      =   84
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
      Begin imText6Ctl.imText txtUrl 
         Height          =   345
         Left            =   2220
         TabIndex        =   30
         Top             =   8220
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
         _ExtentY        =   609
         Caption         =   "frmMt010.frx":2B5A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2BC8
         Key             =   "frmMt010.frx":2BE6
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
         MaxLength       =   50
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
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
         Index           =   24
         Left            =   60
         TabIndex        =   80
         Top             =   8640
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "E-Mail"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   50
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
         Height          =   345
         Left            =   2220
         TabIndex        =   31
         Top             =   8640
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   609
         Caption         =   "frmMt010.frx":2C1A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2C88
         Key             =   "frmMt010.frx":2CA6
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
         Index           =   25
         Left            =   6780
         TabIndex        =   81
         Top             =   1680
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "îÉéÂã£îÑéËêîóø"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   17
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
         Height          =   375
         Left            =   8940
         TabIndex        =   84
         Top             =   1680
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":2CDA
         Caption         =   "frmMt010.frx":2CFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2D68
         Keys            =   "frmMt010.frx":2D86
         Spin            =   "frmMt010.frx":2DD0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1179653
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   26
         Left            =   3360
         TabIndex        =   85
         Top             =   660
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "èoïié“éÛïtì`ï[ë„"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelWidth      =   120
         LabelHeight     =   25
         LabelLeft       =   10
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
         Height          =   375
         Left            =   5520
         TabIndex        =   86
         Top             =   660
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":2DF8
         Caption         =   "frmMt010.frx":2E18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2E86
         Keys            =   "frmMt010.frx":2EA4
         Spin            =   "frmMt010.frx":2EEE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   27
         Left            =   3360
         TabIndex        =   88
         Top             =   1140
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "èoïié“â◊éDë„"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         LabelLeft       =   25
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
         Height          =   375
         Left            =   5520
         TabIndex        =   89
         Top             =   1140
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calculator      =   "frmMt010.frx":2F16
         Caption         =   "frmMt010.frx":2F36
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMt010.frx":2FA4
         Keys            =   "frmMt010.frx":2FC2
         Spin            =   "frmMt010.frx":300C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   -9999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1179653
         Value           =   9999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin VB.Label Label1 
         Caption         =   "â~"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   6360
         TabIndex        =   90
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "â~"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   6360
         TabIndex        =   87
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Åì"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   9780
         TabIndex        =   83
         Top             =   1740
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "ì˙"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   8640
         TabIndex        =   73
         Top             =   3900
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "åé"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   7800
         TabIndex        =   72
         Top             =   3900
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "îN"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   6960
         TabIndex        =   71
         Top             =   3900
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "â~"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   6960
         TabIndex        =   68
         Top             =   3420
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "áu"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   3300
         TabIndex        =   66
         Top             =   5340
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "áu"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3300
         TabIndex        =   64
         Top             =   4860
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "áu"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3300
         TabIndex        =   62
         Top             =   4380
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "áu"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3300
         TabIndex        =   60
         Top             =   3900
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "áu"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3300
         TabIndex        =   58
         Top             =   3420
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "â~"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   9780
         TabIndex        =   56
         Top             =   2700
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "â~"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3060
         TabIndex        =   54
         Top             =   2220
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Åì"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3060
         TabIndex        =   52
         Top             =   2700
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Åì"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   9780
         TabIndex        =   51
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Åì"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3060
         TabIndex        =   50
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "â~"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9780
         TabIndex        =   49
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "â~"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3060
         TabIndex        =   48
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "â~"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9840
      TabIndex        =   82
      Top             =   1740
      Width           =   375
   End
End
Attribute VB_Name = "frmMt010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_clsAdoSQL As New clsAdoCore
Public m_clsReg As New clsReg

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFâÊñ ÉNÉäÉAÉNÉäÉbÉNéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    On Error Resume Next
    txtCompany.SetFocus

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFé¿çsÉNÉäÉbÉNéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Sub cmdExecute_Click()

    On Error Resume Next

    If MsgBox("é¿çsÇµÇ‹Ç∑Ç©ÅH", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    'ì¸óÕÉ`ÉFÉbÉN
    If DoValidationChecks() = False Then Exit Sub
    If DataUpdate() = False Then Exit Sub
    
    txtCompany.SetFocus

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFèIóπÉNÉäÉbÉNéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFÉtÉHÅ[ÉÄÉLÅ[É_ÉEÉìéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Form_KeyDown_Err
    
    'ÉäÉ^Å[ÉìÉLÅ[Ç≈éüÇÃÉRÉìÉgÉçÅ[ÉãÇ÷ÉtÉHÅ[ÉJÉXà⁄ìÆ
    If KeyCode = vbKeyReturn And Shift <> vbShiftMask Then
        KeyCode = 0
        Global_SendKeys Me, VK_TAB
        Exit Sub
    End If
    
    'ÉVÉáÅ[ÉgÉJÉbÉgÉLÅ[ÇÃäÑÇËìñÇƒ
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

    Call MsgBox("ÉtÉHÅ[ÉÄÉLÅ[É_ÉEÉìéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFÉtÉHÅ[ÉÄÉçÅ[Éhéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "ê›íËÉ}ÉXÉ^ï€éÁ"

    'èdï°ãNìÆÇÃÉ`ÉFÉbÉN
    If App.PrevInstance = True Then
        Unload Me
        End
    End If

    'ÉåÉWÉXÉgÉäì«Ç›çûÇ›
    m_clsReg.RegKey = REG_KEY
    If m_clsReg.ReadReg = False Then
        Unload Me
        End
    End If
    
    'ÉfÅ[É^ÉxÅ[ÉXê⁄ë±
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
    
    'ÉtÉBÅ[ÉãÉhÉNÉäÉA
    Call FieldsClear(0)
    
    'ÉtÉBÅ[ÉãÉhÉZÉbÉg
    Call FieldsSet(True)
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("ÉtÉHÅ[ÉÄÉçÅ[ÉhéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    End

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFÉtÉHÅ[ÉÄÉAÉìÉçÅ[Éhéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set m_clsAdoSQL = Nothing
    Set m_clsReg = Nothing
    End
    
    Exit Sub
    
Form_Unload_Err:

    Call MsgBox("ÉtÉHÅ[ÉÄÉAÉìÉçÅ[ÉhéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")
    End

End Sub

'ñ⁄Å@ìIÅ@Å@ÅFâÊñ ÉNÉäÉA
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF0ÅFëSâÊñ 
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        txtCompany.Text = ""
        imnSkeep.Value = 0
        imnBkeep.Value = 0
        imnSrate.Value = 0
        imnBrate.Value = 0
        optSfraction(1).Value = True
        optBfraction(1).Value = True
        imnSRounding.Value = 0
        imnBRounding.Value = 0
        imnCrate.Value = 0
        imnTm2.Value = 0
        imnOm2.Value = 0
        imnIm2.Value = 0
        imnCm2.Value = 0
        imnBm2.Value = 0
        imnUp.Value = 0
        imnPdate(0).Text = ""
        imnPdate(1).Text = ""
        imnPdate(2).Text = ""
        imtMemo.Text = ""
        txtPost.Text = ""
        txtAddres.Text = ""
        txtTel.Text = ""
        txtFax.Text = ""
        txtCeo.Text = ""
        txtUrl.Text = ""
        txtMail.Text = ""
        
        '2011/07
        imnRrate.Value = 0
        imnEf.Value = 0
        imnBrate.Value = 0
        '2011/07
        
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("âÊñ ÉNÉäÉAÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub


Private Sub imnBkeep_GotFocus()

    imnBkeep.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnBkeep_LostFocus()

    imnBkeep.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnBm2_GotFocus()

    imnBm2.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBm2_LostFocus()

    imnBm2.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnBrate_GotFocus()

    imnBrate.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBrate_LostFocus()

    imnBrate.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnBRounding_GotFocus()

    imnBRounding.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBRounding_LostFocus()

    imnBRounding.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnCm2_GotFocus()

    imnCm2.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnCm2_LostFocus()

    imnCm2.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnCrate_GotFocus()

    imnCrate.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnCrate_LostFocus()

    imnCrate.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnIm2_GotFocus()

    imnIm2.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnIm2_LostFocus()

    imnIm2.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnOm2_GotFocus()

    imnOm2.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnOm2_LostFocus()

    imnOm2.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnPdate_GotFocus(Index As Integer)

    imnPdate(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnPdate_LostFocus(Index As Integer)

    imnPdate(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnSkeep_GotFocus()

    imnSkeep.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnSkeep_LostFocus()

    imnSkeep.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnRrate_GotFocus()

    imnRrate.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnRrate_LostFocus()

    imnRrate.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnEf_GotFocus()

    imnEf.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnEf_LostFocus()

    imnEf.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnBrate2_GotFocus()

    imnBrate2.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnBrate2_LostFocus()

    imnBrate2.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnSrate_GotFocus()

    imnSrate.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnSrate_LostFocus()

    imnSrate.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnSRounding_GotFocus()

    imnSRounding.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnSRounding_LostFocus()

    imnSRounding.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnTm2_GotFocus()

    imnTm2.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnTm2_LostFocus()

    imnTm2.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnUp_GotFocus()

    imnUp.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnUp_LostFocus()

    imnUp.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    txtCompany.SetFocus

End Sub

'ñ⁄Å@ìIÅ@Å@ÅFì¸óÕÉ`ÉFÉbÉN
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtCompany.Text) = "" Then
        strErrMsg = "é©é–ñºÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        txtCompany.SetFocus
        GoTo ErrorTrap:
    End If

    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "ì¸óÕÉ`ÉFÉbÉN")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("ì¸óÕÉ`ÉFÉbÉNÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function

'ñ⁄Å@ìIÅ@Å@ÅFÉtÉBÅ[ÉãÉhÇÃÉZÉbÉg
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Public Function FieldsSet(blnVisible As Boolean, Optional adoRecordsetArg As Variant) As Boolean

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    If IsMissing(adoRecordsetArg) = False Then
        Set adoRecordset1 = adoRecordsetArg
    Else
        strSQL = "{call sp_MT010;1}"
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
        txtCompany.Text = Trim(.Fields("Company"))
        imnSkeep.Value = IIf(IsNull(.Fields("Skeep")), 0, .Fields("Skeep"))
        imnBkeep.Value = IIf(IsNull(.Fields("Bkeep")), 0, .Fields("Bkeep"))
        imnSrate.Value = IIf(IsNull(.Fields("Srate")), 0, .Fields("Srate"))
        imnBrate.Value = IIf(IsNull(.Fields("Brate")), 0, .Fields("Brate"))
        If Not IsNull(.Fields("Sfraction")) Then
            optSfraction(.Fields("Sfraction")).Value = True
        End If
        If Not IsNull(.Fields("Bfraction")) Then
            optBfraction(.Fields("Bfraction")).Value = True
        End If
        imnCrate.Value = IIf(IsNull(.Fields("Crate")), 0, .Fields("Crate"))
        imnSRounding.Value = IIf(IsNull(.Fields("SRounding")), 0, .Fields("SRounding"))
        imnBRounding.Value = IIf(IsNull(.Fields("BRounding")), 0, .Fields("BRounding"))
        imnTm2.Value = IIf(IsNull(.Fields("Tm2")), 0, .Fields("Tm2"))
        imnOm2.Value = IIf(IsNull(.Fields("Om2")), 0, .Fields("Om2"))
        imnIm2.Value = IIf(IsNull(.Fields("Im2")), 0, .Fields("Im2"))
        imnCm2.Value = IIf(IsNull(.Fields("Cm2")), 0, .Fields("Cm2"))
        imnBm2.Value = IIf(IsNull(.Fields("Bm2")), 0, .Fields("Bm2"))
        imnUp.Value = IIf(IsNull(.Fields("Up")), 0, .Fields("Up"))
        If Not IsNull(.Fields("Pdate")) Then
            imnPdate(0).Text = IIf(Trim(.Fields("Pdate")) = "", "", left$(.Fields("Pdate"), 4))
            imnPdate(1).Text = CInt(IIf(Trim(.Fields("Pdate")) = "", "", Mid$(.Fields("Pdate"), 6, 2)))
            imnPdate(2).Text = CInt(IIf(Trim(.Fields("Pdate")) = "", "", right$(.Fields("Pdate"), 2)))
        End If
        imtMemo.Text = IIf(IsNull(.Fields("Memo")), "", .Fields("Memo"))
        txtPost.Text = IIf(IsNull(.Fields("CPost")), "", Trim(.Fields("CPost")))
        txtAddres.Text = IIf(IsNull(.Fields("CAddres")), "", Trim(.Fields("CAddres")))
        txtTel.Text = IIf(IsNull(.Fields("CTel")), "", Trim(.Fields("CTel")))
        txtFax.Text = IIf(IsNull(.Fields("CFax")), "", Trim(.Fields("CFax")))
        txtCeo.Text = IIf(IsNull(.Fields("CCeo")), "", Trim(.Fields("CCeo")))
        txtUrl.Text = IIf(IsNull(.Fields("CUrl")), "", Trim(.Fields("CUrl")))
        txtMail.Text = IIf(IsNull(.Fields("CMail")), "", Trim(.Fields("CMail")))
        
        '2011/07
        imnRrate.Value = IIf(IsNull(.Fields("Rrate")), 0, .Fields("Rrate"))
        imnEf.Value = IIf(IsNull(.Fields("Ef")), 0, .Fields("Ef"))
        imnBrate2.Value = IIf(IsNull(.Fields("Brate2")), 0, .Fields("Brate2"))
        '2011/07
        
    End With
    
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
    Screen.MousePointer = vbDefault
    
    FieldsSet = True
    
    Exit Function

FieldsSet_Err:

    FieldsSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("ÉtÉBÅ[ÉãÉhÉZÉbÉgÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsSet_Err")

End Function

'ñ⁄Å@ìIÅ@Å@ÅFÉfÅ[É^ÇÃìoò^
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇPÇO
'çXêVóöóÅ@ÅF
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    m_clsAdoSQL.Connection.BeginTrans
    
    With adoRecordset1
        strSQL = "{call sp_MT010;1}"
        .Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        .Fields("Company") = txtCompany.Text
        .Fields("Skeep") = imnSkeep.Value
        .Fields("Bkeep") = imnBkeep.Value
        .Fields("Srate") = imnSrate.Value
        .Fields("Brate") = imnBrate.Value
        If optSfraction(1).Value = True Then
            .Fields("Sfraction") = FRACTION_OMISSION
        ElseIf optSfraction(2).Value = True Then
            .Fields("Sfraction") = FRACTION_ROUNDOFF
        ElseIf optSfraction(3).Value = True Then
            .Fields("Sfraction") = FRACTION_ROUNDUP
        Else
            .Fields("Sfraction") = FRACTION_OMISSION
        End If
        If optBfraction(1).Value = True Then
            .Fields("Bfraction") = FRACTION_OMISSION
        ElseIf optBfraction(2).Value = True Then
            .Fields("Bfraction") = FRACTION_ROUNDOFF
        ElseIf optBfraction(3).Value = True Then
            .Fields("Bfraction") = FRACTION_ROUNDUP
        Else
            .Fields("Bfraction") = FRACTION_OMISSION
        End If
        .Fields("Crate") = imnCrate.Value
        .Fields("SRounding") = imnSRounding.Value
        .Fields("BRounding") = imnBRounding.Value
        .Fields("Tm2") = imnTm2.Value
        .Fields("Om2") = imnOm2.Value
        .Fields("Im2") = imnIm2.Value
        .Fields("Cm2") = imnCm2.Value
        .Fields("Bm2") = imnBm2.Value
        .Fields("Up") = imnUp.Value
        If Trim(imnPdate(0).Text) <> "" And Trim(imnPdate(1).Text) <> "" And Trim(imnPdate(2).Text) <> "" Then
            .Fields("Pdate") = imnPdate(0).Text & "/" & Format(imnPdate(1).Text, "00") & "/" & Format(imnPdate(2).Text, "00")
        Else
            .Fields("Pdate") = Null
        End If
        .Fields("Memo") = imtMemo.Text
        .Fields("CPost") = txtPost.Text
        .Fields("CAddres") = txtAddres.Text
        .Fields("CTel") = txtTel.Text
        .Fields("CFax") = txtFax.Text
        .Fields("CCeo") = txtCeo.Text
        .Fields("CUrl") = txtUrl.Text
        .Fields("CMail") = txtMail.Text
        
        .Fields("Rrate") = imnRrate.Value
        .Fields("Ef") = imnEf.Value
        .Fields("Brate2") = imnBrate2.Value
        
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
    Call MsgBox("ÉfÅ[É^ìoò^ÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_Err")

End Function

Private Sub imtMemo_GotFocus()

    imtMemo.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imtMemo_LostFocus()

    imtMemo.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub optBfraction_Click(Index As Integer)

    Dim intIndex1 As Integer

    On Error GoTo optBfraction_Click_Err

    'îwåiêFÇÃïœçX
    For intIndex1 = 1 To optBfraction.Count
        If intIndex1 = Index Then
            optBfraction(intIndex1).BackColor = BUTTON_ON
        Else
            optBfraction(intIndex1).BackColor = BUTTON_OFF
        End If
    Next intIndex1

    Exit Sub

optBfraction_Click_Err:

    Call MsgBox("îÉéÂí[êîèàóùÉNÉäÉbÉNéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "optBfraction_Click_Err")

End Sub

Private Sub optSfraction_Click(Index As Integer)

    Dim intIndex1 As Integer

    On Error GoTo optSfraction_Click_Err

    'îwåiêFÇÃïœçX
    For intIndex1 = 1 To optSfraction.Count
        If intIndex1 = Index Then
            optSfraction(intIndex1).BackColor = BUTTON_ON
        Else
            optSfraction(intIndex1).BackColor = BUTTON_OFF
        End If
    Next intIndex1

    Exit Sub

optSfraction_Click_Err:

    Call MsgBox("èoïié“í[êîèàóùÉNÉäÉbÉNéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "optSfraction_Click_Err")

End Sub

Private Sub txtCompany_GotFocus()

    txtCompany.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtCompany_LostFocus()

    txtCompany.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtPost_GotFocus()

   txtPost.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtPost_LostFocus()

   txtPost.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtAddres_GotFocus()

    txtAddres.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtAddres_LostFocus()

    txtAddres.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtTel_GotFocus()
    
    txtTel.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtTel_LostFocus()
    
    txtTel.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtFax_GotFocus()

   txtFax.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtFax_LostFocus()

   txtFax.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtCeo_GotFocus()

   txtCeo.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtCeo_LostFocus()

   txtCeo.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtUrl_GotFocus()

   txtUrl.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtUrl_LostFocus()

   txtUrl.BackColor = FOCUS_NO_COLOR
   
End Sub

Private Sub txtMail_GotFocus()

   txtMail.BackColor = FOCUS_STOP_COLOR
   
End Sub

Private Sub txtMail_LostFocus()

   txtMail.BackColor = FOCUS_NO_COLOR
   
End Sub

