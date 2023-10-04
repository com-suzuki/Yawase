VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmYpmf090 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   ClientHeight    =   10560
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   14880
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf090.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   14880
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   60
      TabIndex        =   177
      Top             =   720
      Width           =   14715
      Begin VB.CommandButton cmdSearch 
         Caption         =   "åüçıäJén"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12420
         TabIndex        =   3
         Top             =   180
         Width           =   1635
      End
      Begin CSComboLib.CSComboBox cboPnum 
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
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf090.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   178
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "éÛïtî‘çÜ"
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
      Begin CSComboLib.CSComboBox cboPnum 
         Height          =   405
         Index           =   1
         Left            =   7080
         TabIndex        =   2
         Top             =   180
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf090.frx":0D13
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin VB.Label lblPnum_Name 
         Appearance      =   0  'Ã◊Øƒ
         BackColor       =   &H80000005&
         BorderStyle     =   1  'é¿ê¸
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         TabIndex        =   181
         Top             =   180
         Width           =   3855
      End
      Begin VB.Label lblPnum_Name 
         Appearance      =   0  'Ã◊Øƒ
         BackColor       =   &H80000005&
         BorderStyle     =   1  'é¿ê¸
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   8160
         TabIndex        =   180
         Top             =   180
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "Å`"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   179
         Top             =   180
         Width           =   435
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9960
      Top             =   120
   End
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton cmdLogin 
         Caption         =   "äJç√îNåéì˙Ç∆íSìñé“ÇÃïœçX"
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
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "2;20"
         Contents        =   "frmYpmf090.frx":0D2C
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
         Caption         =   "äJç√îNåéì˙"
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
         Caption         =   "íSìñé“"
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
         Appearance      =   0  'Ã◊Øƒ
         BackColor       =   &H80000005&
         BorderStyle     =   1  'é¿ê¸
         Caption         =   "ÇmÇmÇmÇmÇmÇmÇmÇmÇmÇm"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Alignment       =   2  'íÜâõëµÇ¶
         Appearance      =   0  'Ã◊Øƒ
         BackColor       =   &H80000005&
         BorderStyle     =   1  'é¿ê¸
         Caption         =   "9999/12/31"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Top             =   9720
      Width           =   14715
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
         Caption         =   "çƒï\é¶(F8)"
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
         rText.left      =   15
         rText.top       =   8
         rText.right     =   97
         rText.bottom    =   27
         Picture         =   "frmYpmf090.frx":0D45
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
         Picture         =   "frmYpmf090.frx":0D61
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   10860
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "àÛç¸(F12)"
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
         rPic.left       =   10
         rPic.top        =   8
         rPic.right      =   16
         rPic.bottom     =   16
         rText.left      =   30
         rText.top       =   8
         rText.right     =   105
         rText.bottom    =   27
         Picture         =   "frmYpmf090.frx":0EBB
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   7575
      Left            =   60
      TabIndex        =   9
      Top             =   1560
      Width           =   14715
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   107
         Top             =   6780
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   10
            Left            =   4800
            Picture         =   "frmYpmf090.frx":0FCD
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   201
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   10
            Left            =   5220
            Picture         =   "frmYpmf090.frx":12D7
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   191
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   10
            Left            =   3420
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":13D9
            Caption         =   "frmYpmf090.frx":13F9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":1467
            Keys            =   "frmYpmf090.frx":1485
            Spin            =   "frmYpmf090.frx":14CF
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   10
            Left            =   5640
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":14F7
            Caption         =   "frmYpmf090.frx":1517
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":1585
            Keys            =   "frmYpmf090.frx":15A3
            Spin            =   "frmYpmf090.frx":15ED
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
            Left            =   10915
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":1615
            Caption         =   "frmYpmf090.frx":1635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":16A3
            Keys            =   "frmYpmf090.frx":16C1
            Spin            =   "frmYpmf090.frx":170B
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
            Left            =   9700
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":1733
            Caption         =   "frmYpmf090.frx":1753
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":17C1
            Keys            =   "frmYpmf090.frx":17DF
            Spin            =   "frmYpmf090.frx":1829
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
            Left            =   7860
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":1851
            Caption         =   "frmYpmf090.frx":1871
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":18DF
            Keys            =   "frmYpmf090.frx":18FD
            Spin            =   "frmYpmf090.frx":1947
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
            Left            =   12450
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":196F
            Caption         =   "frmYpmf090.frx":198F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":19FD
            Keys            =   "frmYpmf090.frx":1A1B
            Spin            =   "frmYpmf090.frx":1A65
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   10
            Left            =   6850
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":1A8D
            Caption         =   "frmYpmf090.frx":1AAD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":1B1B
            Keys            =   "frmYpmf090.frx":1B39
            Spin            =   "frmYpmf090.frx":1B83
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   10
            Left            =   4380
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":1BAB
            Caption         =   "frmYpmf090.frx":1BCB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":1C39
            Keys            =   "frmYpmf090.frx":1C57
            Spin            =   "frmYpmf090.frx":1CA1
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   10
            Left            =   3840
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":1CC9
            Caption         =   "frmYpmf090.frx":1CE9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":1D57
            Keys            =   "frmYpmf090.frx":1D75
            Spin            =   "frmYpmf090.frx":1DBF
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   10
            Left            =   8470
            TabIndex        =   232
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":1DE7
            Caption         =   "frmYpmf090.frx":1E07
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":1E75
            Keys            =   "frmYpmf090.frx":1E93
            Spin            =   "frmYpmf090.frx":1EDD
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   10
            Left            =   9090
            TabIndex        =   233
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":1F05
            Caption         =   "frmYpmf090.frx":1F25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":1F93
            Keys            =   "frmYpmf090.frx":1FB1
            Spin            =   "frmYpmf090.frx":1FFB
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   10
            Left            =   11840
            TabIndex        =   234
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":2023
            Caption         =   "frmYpmf090.frx":2043
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":20B1
            Keys            =   "frmYpmf090.frx":20CF
            Spin            =   "frmYpmf090.frx":2119
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
            ValueVT         =   1179653
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   176
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   125
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   115
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
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
         Top             =   6120
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   9
            Left            =   4800
            Picture         =   "frmYpmf090.frx":2141
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   200
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   9
            Left            =   5220
            Picture         =   "frmYpmf090.frx":244B
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   190
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   9
            Left            =   3420
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":254D
            Caption         =   "frmYpmf090.frx":256D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":25DB
            Keys            =   "frmYpmf090.frx":25F9
            Spin            =   "frmYpmf090.frx":2643
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   9
            Left            =   5640
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":266B
            Caption         =   "frmYpmf090.frx":268B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":26F9
            Keys            =   "frmYpmf090.frx":2717
            Spin            =   "frmYpmf090.frx":2761
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
            Left            =   10915
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":2789
            Caption         =   "frmYpmf090.frx":27A9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":2817
            Keys            =   "frmYpmf090.frx":2835
            Spin            =   "frmYpmf090.frx":287F
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
            Left            =   9700
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":28A7
            Caption         =   "frmYpmf090.frx":28C7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":2935
            Keys            =   "frmYpmf090.frx":2953
            Spin            =   "frmYpmf090.frx":299D
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
            Left            =   7860
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":29C5
            Caption         =   "frmYpmf090.frx":29E5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":2A53
            Keys            =   "frmYpmf090.frx":2A71
            Spin            =   "frmYpmf090.frx":2ABB
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
            Left            =   12450
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":2AE3
            Caption         =   "frmYpmf090.frx":2B03
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":2B71
            Keys            =   "frmYpmf090.frx":2B8F
            Spin            =   "frmYpmf090.frx":2BD9
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   9
            Left            =   6850
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":2C01
            Caption         =   "frmYpmf090.frx":2C21
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":2C8F
            Keys            =   "frmYpmf090.frx":2CAD
            Spin            =   "frmYpmf090.frx":2CF7
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   9
            Left            =   4380
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":2D1F
            Caption         =   "frmYpmf090.frx":2D3F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":2DAD
            Keys            =   "frmYpmf090.frx":2DCB
            Spin            =   "frmYpmf090.frx":2E15
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   9
            Left            =   3840
            TabIndex        =   172
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":2E3D
            Caption         =   "frmYpmf090.frx":2E5D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":2ECB
            Keys            =   "frmYpmf090.frx":2EE9
            Spin            =   "frmYpmf090.frx":2F33
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
            ValueVT         =   5
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   9
            Left            =   8470
            TabIndex        =   229
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":2F5B
            Caption         =   "frmYpmf090.frx":2F7B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":2FE9
            Keys            =   "frmYpmf090.frx":3007
            Spin            =   "frmYpmf090.frx":3051
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   9
            Left            =   9090
            TabIndex        =   230
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":3079
            Caption         =   "frmYpmf090.frx":3099
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":3107
            Keys            =   "frmYpmf090.frx":3125
            Spin            =   "frmYpmf090.frx":316F
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   9
            Left            =   11840
            TabIndex        =   231
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":3197
            Caption         =   "frmYpmf090.frx":31B7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":3225
            Keys            =   "frmYpmf090.frx":3243
            Spin            =   "frmYpmf090.frx":328D
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
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   173
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   124
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
            TabIndex        =   106
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   105
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   89
         Top             =   5460
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   8
            Left            =   4800
            Picture         =   "frmYpmf090.frx":32B5
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   199
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   8
            Left            =   5220
            Picture         =   "frmYpmf090.frx":35BF
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   8
            Left            =   3420
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":36C1
            Caption         =   "frmYpmf090.frx":36E1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":374F
            Keys            =   "frmYpmf090.frx":376D
            Spin            =   "frmYpmf090.frx":37B7
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   8
            Left            =   5640
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":37DF
            Caption         =   "frmYpmf090.frx":37FF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":386D
            Keys            =   "frmYpmf090.frx":388B
            Spin            =   "frmYpmf090.frx":38D5
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
            Left            =   10915
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":38FD
            Caption         =   "frmYpmf090.frx":391D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":398B
            Keys            =   "frmYpmf090.frx":39A9
            Spin            =   "frmYpmf090.frx":39F3
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
            Left            =   9700
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":3A1B
            Caption         =   "frmYpmf090.frx":3A3B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":3AA9
            Keys            =   "frmYpmf090.frx":3AC7
            Spin            =   "frmYpmf090.frx":3B11
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
            Left            =   7860
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":3B39
            Caption         =   "frmYpmf090.frx":3B59
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":3BC7
            Keys            =   "frmYpmf090.frx":3BE5
            Spin            =   "frmYpmf090.frx":3C2F
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
            Left            =   12450
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":3C57
            Caption         =   "frmYpmf090.frx":3C77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":3CE5
            Keys            =   "frmYpmf090.frx":3D03
            Spin            =   "frmYpmf090.frx":3D4D
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   8
            Left            =   6850
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":3D75
            Caption         =   "frmYpmf090.frx":3D95
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":3E03
            Keys            =   "frmYpmf090.frx":3E21
            Spin            =   "frmYpmf090.frx":3E6B
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   8
            Left            =   4380
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":3E93
            Caption         =   "frmYpmf090.frx":3EB3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":3F21
            Keys            =   "frmYpmf090.frx":3F3F
            Spin            =   "frmYpmf090.frx":3F89
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   8
            Left            =   3840
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":3FB1
            Caption         =   "frmYpmf090.frx":3FD1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":403F
            Keys            =   "frmYpmf090.frx":405D
            Spin            =   "frmYpmf090.frx":40A7
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   8
            Left            =   8470
            TabIndex        =   226
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":40CF
            Caption         =   "frmYpmf090.frx":40EF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":415D
            Keys            =   "frmYpmf090.frx":417B
            Spin            =   "frmYpmf090.frx":41C5
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   8
            Left            =   9090
            TabIndex        =   227
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":41ED
            Caption         =   "frmYpmf090.frx":420D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":427B
            Keys            =   "frmYpmf090.frx":4299
            Spin            =   "frmYpmf090.frx":42E3
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   8
            Left            =   11840
            TabIndex        =   228
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":430B
            Caption         =   "frmYpmf090.frx":432B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":4399
            Keys            =   "frmYpmf090.frx":43B7
            Spin            =   "frmYpmf090.frx":4401
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
            ValueVT         =   1179653
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   170
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   123
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
            TabIndex        =   97
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   96
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   80
         Top             =   4800
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   7
            Left            =   4800
            Picture         =   "frmYpmf090.frx":4429
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   198
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   7
            Left            =   5220
            Picture         =   "frmYpmf090.frx":4733
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   7
            Left            =   3420
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":4835
            Caption         =   "frmYpmf090.frx":4855
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":48C3
            Keys            =   "frmYpmf090.frx":48E1
            Spin            =   "frmYpmf090.frx":492B
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   7
            Left            =   5640
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":4953
            Caption         =   "frmYpmf090.frx":4973
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":49E1
            Keys            =   "frmYpmf090.frx":49FF
            Spin            =   "frmYpmf090.frx":4A49
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
            Left            =   10915
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":4A71
            Caption         =   "frmYpmf090.frx":4A91
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":4AFF
            Keys            =   "frmYpmf090.frx":4B1D
            Spin            =   "frmYpmf090.frx":4B67
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
            Left            =   9700
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":4B8F
            Caption         =   "frmYpmf090.frx":4BAF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":4C1D
            Keys            =   "frmYpmf090.frx":4C3B
            Spin            =   "frmYpmf090.frx":4C85
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
            Left            =   7860
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":4CAD
            Caption         =   "frmYpmf090.frx":4CCD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":4D3B
            Keys            =   "frmYpmf090.frx":4D59
            Spin            =   "frmYpmf090.frx":4DA3
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
            Left            =   12450
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":4DCB
            Caption         =   "frmYpmf090.frx":4DEB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":4E59
            Keys            =   "frmYpmf090.frx":4E77
            Spin            =   "frmYpmf090.frx":4EC1
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   7
            Left            =   6850
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":4EE9
            Caption         =   "frmYpmf090.frx":4F09
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":4F77
            Keys            =   "frmYpmf090.frx":4F95
            Spin            =   "frmYpmf090.frx":4FDF
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   7
            Left            =   4380
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5007
            Caption         =   "frmYpmf090.frx":5027
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":5095
            Keys            =   "frmYpmf090.frx":50B3
            Spin            =   "frmYpmf090.frx":50FD
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   7
            Left            =   3840
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5125
            Caption         =   "frmYpmf090.frx":5145
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":51B3
            Keys            =   "frmYpmf090.frx":51D1
            Spin            =   "frmYpmf090.frx":521B
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   7
            Left            =   8470
            TabIndex        =   223
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5243
            Caption         =   "frmYpmf090.frx":5263
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":52D1
            Keys            =   "frmYpmf090.frx":52EF
            Spin            =   "frmYpmf090.frx":5339
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   7
            Left            =   9090
            TabIndex        =   224
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5361
            Caption         =   "frmYpmf090.frx":5381
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":53EF
            Keys            =   "frmYpmf090.frx":540D
            Spin            =   "frmYpmf090.frx":5457
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   7
            Left            =   11840
            TabIndex        =   225
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":547F
            Caption         =   "frmYpmf090.frx":549F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":550D
            Keys            =   "frmYpmf090.frx":552B
            Spin            =   "frmYpmf090.frx":5575
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
            ValueVT         =   1179653
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   167
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   122
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
            TabIndex        =   88
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   87
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   71
         Top             =   4140
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   6
            Left            =   4800
            Picture         =   "frmYpmf090.frx":559D
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   197
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   6
            Left            =   5220
            Picture         =   "frmYpmf090.frx":58A7
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   187
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   6
            Left            =   3420
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":59A9
            Caption         =   "frmYpmf090.frx":59C9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":5A37
            Keys            =   "frmYpmf090.frx":5A55
            Spin            =   "frmYpmf090.frx":5A9F
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   6
            Left            =   5640
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5AC7
            Caption         =   "frmYpmf090.frx":5AE7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":5B55
            Keys            =   "frmYpmf090.frx":5B73
            Spin            =   "frmYpmf090.frx":5BBD
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
            Left            =   10915
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5BE5
            Caption         =   "frmYpmf090.frx":5C05
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":5C73
            Keys            =   "frmYpmf090.frx":5C91
            Spin            =   "frmYpmf090.frx":5CDB
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
            Left            =   9700
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5D03
            Caption         =   "frmYpmf090.frx":5D23
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":5D91
            Keys            =   "frmYpmf090.frx":5DAF
            Spin            =   "frmYpmf090.frx":5DF9
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
            Left            =   7860
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5E21
            Caption         =   "frmYpmf090.frx":5E41
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":5EAF
            Keys            =   "frmYpmf090.frx":5ECD
            Spin            =   "frmYpmf090.frx":5F17
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
            Left            =   12450
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":5F3F
            Caption         =   "frmYpmf090.frx":5F5F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":5FCD
            Keys            =   "frmYpmf090.frx":5FEB
            Spin            =   "frmYpmf090.frx":6035
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   6
            Left            =   6850
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":605D
            Caption         =   "frmYpmf090.frx":607D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":60EB
            Keys            =   "frmYpmf090.frx":6109
            Spin            =   "frmYpmf090.frx":6153
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   6
            Left            =   4380
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":617B
            Caption         =   "frmYpmf090.frx":619B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6209
            Keys            =   "frmYpmf090.frx":6227
            Spin            =   "frmYpmf090.frx":6271
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   6
            Left            =   3840
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":6299
            Caption         =   "frmYpmf090.frx":62B9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6327
            Keys            =   "frmYpmf090.frx":6345
            Spin            =   "frmYpmf090.frx":638F
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   6
            Left            =   8470
            TabIndex        =   220
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":63B7
            Caption         =   "frmYpmf090.frx":63D7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6445
            Keys            =   "frmYpmf090.frx":6463
            Spin            =   "frmYpmf090.frx":64AD
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   6
            Left            =   9090
            TabIndex        =   221
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":64D5
            Caption         =   "frmYpmf090.frx":64F5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6563
            Keys            =   "frmYpmf090.frx":6581
            Spin            =   "frmYpmf090.frx":65CB
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   6
            Left            =   11840
            TabIndex        =   222
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":65F3
            Caption         =   "frmYpmf090.frx":6613
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6681
            Keys            =   "frmYpmf090.frx":669F
            Spin            =   "frmYpmf090.frx":66E9
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
            ValueVT         =   1179653
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   164
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   121
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
            TabIndex        =   79
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   78
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Top             =   3480
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   5
            Left            =   4800
            Picture         =   "frmYpmf090.frx":6711
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   5
            Left            =   5220
            Picture         =   "frmYpmf090.frx":6A1B
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   5
            Left            =   3420
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":6B1D
            Caption         =   "frmYpmf090.frx":6B3D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6BAB
            Keys            =   "frmYpmf090.frx":6BC9
            Spin            =   "frmYpmf090.frx":6C13
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   5
            Left            =   5640
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":6C3B
            Caption         =   "frmYpmf090.frx":6C5B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6CC9
            Keys            =   "frmYpmf090.frx":6CE7
            Spin            =   "frmYpmf090.frx":6D31
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
            Left            =   10915
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":6D59
            Caption         =   "frmYpmf090.frx":6D79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6DE7
            Keys            =   "frmYpmf090.frx":6E05
            Spin            =   "frmYpmf090.frx":6E4F
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
            Left            =   9700
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":6E77
            Caption         =   "frmYpmf090.frx":6E97
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":6F05
            Keys            =   "frmYpmf090.frx":6F23
            Spin            =   "frmYpmf090.frx":6F6D
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
            Left            =   7860
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":6F95
            Caption         =   "frmYpmf090.frx":6FB5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":7023
            Keys            =   "frmYpmf090.frx":7041
            Spin            =   "frmYpmf090.frx":708B
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
            Left            =   12450
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":70B3
            Caption         =   "frmYpmf090.frx":70D3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":7141
            Keys            =   "frmYpmf090.frx":715F
            Spin            =   "frmYpmf090.frx":71A9
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   5
            Left            =   6850
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":71D1
            Caption         =   "frmYpmf090.frx":71F1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":725F
            Keys            =   "frmYpmf090.frx":727D
            Spin            =   "frmYpmf090.frx":72C7
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   5
            Left            =   4380
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":72EF
            Caption         =   "frmYpmf090.frx":730F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":737D
            Keys            =   "frmYpmf090.frx":739B
            Spin            =   "frmYpmf090.frx":73E5
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   5
            Left            =   3840
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":740D
            Caption         =   "frmYpmf090.frx":742D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":749B
            Keys            =   "frmYpmf090.frx":74B9
            Spin            =   "frmYpmf090.frx":7503
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   5
            Left            =   8470
            TabIndex        =   217
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":752B
            Caption         =   "frmYpmf090.frx":754B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":75B9
            Keys            =   "frmYpmf090.frx":75D7
            Spin            =   "frmYpmf090.frx":7621
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   5
            Left            =   9090
            TabIndex        =   218
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":7649
            Caption         =   "frmYpmf090.frx":7669
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":76D7
            Keys            =   "frmYpmf090.frx":76F5
            Spin            =   "frmYpmf090.frx":773F
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   5
            Left            =   11840
            TabIndex        =   219
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":7767
            Caption         =   "frmYpmf090.frx":7787
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":77F5
            Keys            =   "frmYpmf090.frx":7813
            Spin            =   "frmYpmf090.frx":785D
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
            ValueVT         =   1179653
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   161
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   120
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   70
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
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
         Top             =   2820
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   4
            Left            =   4800
            Picture         =   "frmYpmf090.frx":7885
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   195
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   4
            Left            =   5220
            Picture         =   "frmYpmf090.frx":7B8F
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   185
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   4
            Left            =   3420
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":7C91
            Caption         =   "frmYpmf090.frx":7CB1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":7D1F
            Keys            =   "frmYpmf090.frx":7D3D
            Spin            =   "frmYpmf090.frx":7D87
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   4
            Left            =   5640
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":7DAF
            Caption         =   "frmYpmf090.frx":7DCF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":7E3D
            Keys            =   "frmYpmf090.frx":7E5B
            Spin            =   "frmYpmf090.frx":7EA5
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
            Left            =   10915
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":7ECD
            Caption         =   "frmYpmf090.frx":7EED
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":7F5B
            Keys            =   "frmYpmf090.frx":7F79
            Spin            =   "frmYpmf090.frx":7FC3
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
            Left            =   9700
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":7FEB
            Caption         =   "frmYpmf090.frx":800B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":8079
            Keys            =   "frmYpmf090.frx":8097
            Spin            =   "frmYpmf090.frx":80E1
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
            Left            =   7860
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":8109
            Caption         =   "frmYpmf090.frx":8129
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":8197
            Keys            =   "frmYpmf090.frx":81B5
            Spin            =   "frmYpmf090.frx":81FF
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
            Left            =   12450
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":8227
            Caption         =   "frmYpmf090.frx":8247
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":82B5
            Keys            =   "frmYpmf090.frx":82D3
            Spin            =   "frmYpmf090.frx":831D
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   4
            Left            =   6850
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":8345
            Caption         =   "frmYpmf090.frx":8365
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":83D3
            Keys            =   "frmYpmf090.frx":83F1
            Spin            =   "frmYpmf090.frx":843B
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   4
            Left            =   4380
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":8463
            Caption         =   "frmYpmf090.frx":8483
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":84F1
            Keys            =   "frmYpmf090.frx":850F
            Spin            =   "frmYpmf090.frx":8559
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   4
            Left            =   3840
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":8581
            Caption         =   "frmYpmf090.frx":85A1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":860F
            Keys            =   "frmYpmf090.frx":862D
            Spin            =   "frmYpmf090.frx":8677
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   4
            Left            =   8470
            TabIndex        =   214
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":869F
            Caption         =   "frmYpmf090.frx":86BF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":872D
            Keys            =   "frmYpmf090.frx":874B
            Spin            =   "frmYpmf090.frx":8795
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   4
            Left            =   9090
            TabIndex        =   215
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":87BD
            Caption         =   "frmYpmf090.frx":87DD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":884B
            Keys            =   "frmYpmf090.frx":8869
            Spin            =   "frmYpmf090.frx":88B3
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   4
            Left            =   11840
            TabIndex        =   216
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":88DB
            Caption         =   "frmYpmf090.frx":88FB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":8969
            Keys            =   "frmYpmf090.frx":8987
            Spin            =   "frmYpmf090.frx":89D1
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
            ValueVT         =   1179653
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   158
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   119
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   61
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
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
         Top             =   2160
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   3
            Left            =   4800
            Picture         =   "frmYpmf090.frx":89F9
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   3
            Left            =   5220
            Picture         =   "frmYpmf090.frx":8D03
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   184
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   3
            Left            =   3420
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":8E05
            Caption         =   "frmYpmf090.frx":8E25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":8E93
            Keys            =   "frmYpmf090.frx":8EB1
            Spin            =   "frmYpmf090.frx":8EFB
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   3
            Left            =   5640
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":8F23
            Caption         =   "frmYpmf090.frx":8F43
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":8FB1
            Keys            =   "frmYpmf090.frx":8FCF
            Spin            =   "frmYpmf090.frx":9019
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
            Left            =   10915
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":9041
            Caption         =   "frmYpmf090.frx":9061
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":90CF
            Keys            =   "frmYpmf090.frx":90ED
            Spin            =   "frmYpmf090.frx":9137
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
            Left            =   9700
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":915F
            Caption         =   "frmYpmf090.frx":917F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":91ED
            Keys            =   "frmYpmf090.frx":920B
            Spin            =   "frmYpmf090.frx":9255
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
            Left            =   7860
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":927D
            Caption         =   "frmYpmf090.frx":929D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":930B
            Keys            =   "frmYpmf090.frx":9329
            Spin            =   "frmYpmf090.frx":9373
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
            Left            =   12450
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":939B
            Caption         =   "frmYpmf090.frx":93BB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":9429
            Keys            =   "frmYpmf090.frx":9447
            Spin            =   "frmYpmf090.frx":9491
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   3
            Left            =   6850
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":94B9
            Caption         =   "frmYpmf090.frx":94D9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":9547
            Keys            =   "frmYpmf090.frx":9565
            Spin            =   "frmYpmf090.frx":95AF
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   3
            Left            =   4380
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":95D7
            Caption         =   "frmYpmf090.frx":95F7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":9665
            Keys            =   "frmYpmf090.frx":9683
            Spin            =   "frmYpmf090.frx":96CD
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   3
            Left            =   3840
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":96F5
            Caption         =   "frmYpmf090.frx":9715
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":9783
            Keys            =   "frmYpmf090.frx":97A1
            Spin            =   "frmYpmf090.frx":97EB
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   3
            Left            =   8470
            TabIndex        =   211
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":9813
            Caption         =   "frmYpmf090.frx":9833
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":98A1
            Keys            =   "frmYpmf090.frx":98BF
            Spin            =   "frmYpmf090.frx":9909
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   3
            Left            =   9090
            TabIndex        =   212
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":9931
            Caption         =   "frmYpmf090.frx":9951
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":99BF
            Keys            =   "frmYpmf090.frx":99DD
            Spin            =   "frmYpmf090.frx":9A27
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   3
            Left            =   11840
            TabIndex        =   213
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":9A4F
            Caption         =   "frmYpmf090.frx":9A6F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":9ADD
            Keys            =   "frmYpmf090.frx":9AFB
            Spin            =   "frmYpmf090.frx":9B45
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
            ValueVT         =   1179653
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   155
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   118
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   52
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
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
         Top             =   1500
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   2
            Left            =   4800
            Picture         =   "frmYpmf090.frx":9B6D
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   2
            Left            =   5220
            Picture         =   "frmYpmf090.frx":9E77
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   2
            Left            =   3420
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":9F79
            Caption         =   "frmYpmf090.frx":9F99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A007
            Keys            =   "frmYpmf090.frx":A025
            Spin            =   "frmYpmf090.frx":A06F
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   2
            Left            =   5640
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A097
            Caption         =   "frmYpmf090.frx":A0B7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A125
            Keys            =   "frmYpmf090.frx":A143
            Spin            =   "frmYpmf090.frx":A18D
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
            Left            =   10915
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A1B5
            Caption         =   "frmYpmf090.frx":A1D5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A243
            Keys            =   "frmYpmf090.frx":A261
            Spin            =   "frmYpmf090.frx":A2AB
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
            Left            =   9700
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A2D3
            Caption         =   "frmYpmf090.frx":A2F3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A361
            Keys            =   "frmYpmf090.frx":A37F
            Spin            =   "frmYpmf090.frx":A3C9
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
            Left            =   7860
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A3F1
            Caption         =   "frmYpmf090.frx":A411
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A47F
            Keys            =   "frmYpmf090.frx":A49D
            Spin            =   "frmYpmf090.frx":A4E7
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
            Left            =   12450
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A50F
            Caption         =   "frmYpmf090.frx":A52F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A59D
            Keys            =   "frmYpmf090.frx":A5BB
            Spin            =   "frmYpmf090.frx":A605
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   2
            Left            =   6850
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   180
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A62D
            Caption         =   "frmYpmf090.frx":A64D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A6BB
            Keys            =   "frmYpmf090.frx":A6D9
            Spin            =   "frmYpmf090.frx":A723
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   2
            Left            =   4380
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A74B
            Caption         =   "frmYpmf090.frx":A76B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A7D9
            Keys            =   "frmYpmf090.frx":A7F7
            Spin            =   "frmYpmf090.frx":A841
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   2
            Left            =   3840
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A869
            Caption         =   "frmYpmf090.frx":A889
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":A8F7
            Keys            =   "frmYpmf090.frx":A915
            Spin            =   "frmYpmf090.frx":A95F
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   2
            Left            =   8470
            TabIndex        =   208
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":A987
            Caption         =   "frmYpmf090.frx":A9A7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":AA15
            Keys            =   "frmYpmf090.frx":AA33
            Spin            =   "frmYpmf090.frx":AA7D
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   2
            Left            =   9090
            TabIndex        =   209
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":AAA5
            Caption         =   "frmYpmf090.frx":AAC5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":AB33
            Keys            =   "frmYpmf090.frx":AB51
            Spin            =   "frmYpmf090.frx":AB9B
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   2
            Left            =   11840
            TabIndex        =   210
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":ABC3
            Caption         =   "frmYpmf090.frx":ABE3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":AC51
            Keys            =   "frmYpmf090.frx":AC6F
            Spin            =   "frmYpmf090.frx":ACB9
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
            ValueVT         =   1179653
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   152
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   117
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   43
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
            TabIndex        =   42
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6495
         Left            =   14280
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   960
         Width           =   315
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Width           =   14115
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "àÛéÜ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   11760
            TabIndex        =   207
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "â◊éD"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   9120
            TabIndex        =   205
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "éÛì`"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   8470
            TabIndex        =   204
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "èÛë‘"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   3840
            TabIndex        =   149
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "éËêîóø"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   6900
            TabIndex        =   145
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "âÒ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   3420
            TabIndex        =   132
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "ëççáåv"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   12480
            TabIndex        =   33
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "à€éù"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   7920
            TabIndex        =   32
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "ç∑à¯ã‡äz"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   9700
            TabIndex        =   31
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "è¡îÔê≈"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   10910
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "îÑóßçáåv"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   5700
            TabIndex        =   29
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "éÛïtî‘çÜ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Begin VB.Frame fraDetail 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   14115
         Begin VB.CommandButton cmdF_Search 
            Height          =   375
            Index           =   1
            Left            =   4800
            Picture         =   "frmYpmf090.frx":ACE1
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   192
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Index           =   1
            Left            =   5220
            Picture         =   "frmYpmf090.frx":AFEB
            Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
            TabIndex        =   182
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
         End
         Begin imNumber6Ctl.imNumber imnNum 
            Height          =   375
            Index           =   1
            Left            =   3420
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B0ED
            Caption         =   "frmYpmf090.frx":B10D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":B17B
            Keys            =   "frmYpmf090.frx":B199
            Spin            =   "frmYpmf090.frx":B1E3
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
            ValueVT         =   2088828933
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnTotal 
            Height          =   375
            Index           =   1
            Left            =   5640
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B20B
            Caption         =   "frmYpmf090.frx":B22B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":B299
            Keys            =   "frmYpmf090.frx":B2B7
            Spin            =   "frmYpmf090.frx":B301
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
            Left            =   10915
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   180
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B329
            Caption         =   "frmYpmf090.frx":B349
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":B3B7
            Keys            =   "frmYpmf090.frx":B3D5
            Spin            =   "frmYpmf090.frx":B41F
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
            Left            =   9700
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B447
            Caption         =   "frmYpmf090.frx":B467
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":B4D5
            Keys            =   "frmYpmf090.frx":B4F3
            Spin            =   "frmYpmf090.frx":B53D
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
            Left            =   7860
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B565
            Caption         =   "frmYpmf090.frx":B585
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":B5F3
            Keys            =   "frmYpmf090.frx":B611
            Spin            =   "frmYpmf090.frx":B65B
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
            Left            =   12450
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B683
            Caption         =   "frmYpmf090.frx":B6A3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":B711
            Keys            =   "frmYpmf090.frx":B72F
            Spin            =   "frmYpmf090.frx":B779
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
         Begin imNumber6Ctl.imNumber imnCharge 
            Height          =   375
            Index           =   1
            Left            =   6850
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   180
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B7A1
            Caption         =   "frmYpmf090.frx":B7C1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":B82F
            Keys            =   "frmYpmf090.frx":B84D
            Spin            =   "frmYpmf090.frx":B897
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
         Begin imNumber6Ctl.imNumber imnFtotal 
            Height          =   375
            Index           =   1
            Left            =   4380
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B8BF
            Caption         =   "frmYpmf090.frx":B8DF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":B94D
            Keys            =   "frmYpmf090.frx":B96B
            Spin            =   "frmYpmf090.frx":B9B5
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnF 
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   147
            TabStop         =   0   'False
            Top             =   180
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":B9DD
            Caption         =   "frmYpmf090.frx":B9FD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":BA6B
            Keys            =   "frmYpmf090.frx":BA89
            Spin            =   "frmYpmf090.frx":BAD3
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
            ValueVT         =   1245189
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRrate 
            Height          =   375
            Index           =   1
            Left            =   8470
            TabIndex        =   202
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":BAFB
            Caption         =   "frmYpmf090.frx":BB1B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":BB89
            Keys            =   "frmYpmf090.frx":BBA7
            Spin            =   "frmYpmf090.frx":BBF1
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
         Begin imNumber6Ctl.imNumber imnEf 
            Height          =   375
            Index           =   1
            Left            =   9090
            TabIndex        =   203
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":BC19
            Caption         =   "frmYpmf090.frx":BC39
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":BCA7
            Keys            =   "frmYpmf090.frx":BCC5
            Spin            =   "frmYpmf090.frx":BD0F
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
         Begin imNumber6Ctl.imNumber imnStamp 
            Height          =   375
            Index           =   1
            Left            =   11840
            TabIndex        =   206
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            _Version        =   65536
            _ExtentX        =   1058
            _ExtentY        =   661
            Calculator      =   "frmYpmf090.frx":BD37
            Caption         =   "frmYpmf090.frx":BD57
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf090.frx":BDC5
            Keys            =   "frmYpmf090.frx":BDE3
            Spin            =   "frmYpmf090.frx":BE2D
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
         Begin VB.Label Label1 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BackStyle       =   0  'ìßñæ
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   4200
            TabIndex        =   148
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblDiv 
            Alignment       =   2  'íÜâõëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "çœ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   13670
            TabIndex        =   116
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblPnumName 
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   720
            TabIndex        =   20
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label lblPnum 
            Alignment       =   1  'âEëµÇ¶
            Appearance      =   0  'Ã◊Øƒ
            BackColor       =   &H80000005&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Left            =   60
            TabIndex        =   19
            Top             =   180
            Width           =   615
         End
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   15060
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf090.frx":BE55
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf090.frx":BEC3
      Key             =   "frmYpmf090.frx":BEE1
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
      Left            =   15060
      TabIndex        =   7
      Top             =   180
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf090.frx":BF25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf090.frx":BF93
      Key             =   "frmYpmf090.frx":BFB1
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
      Left            =   15120
      TabIndex        =   8
      Top             =   1200
      Width           =   90
      _Version        =   65536
      _ExtentX        =   159
      _ExtentY        =   132
      Caption         =   "frmYpmf090.frx":BFF5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf090.frx":C063
      Key             =   "frmYpmf090.frx":C081
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
      Left            =   12300
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   120
      Width           =   2115
      _Version        =   262145
      _ExtentX        =   3731
      _ExtentY        =   767
      _StockProps     =   15
      Caption         =   "èoïié“ê∏éZëSâèú"
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
      rText.left      =   4
      rText.top       =   6
      rText.right     =   136
      rText.bottom    =   25
      Picture         =   "frmYpmf090.frx":C0C5
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   126
      Top             =   9120
      Width           =   14715
      Begin imNumber6Ctl.imNumber imnTotal_Total 
         Height          =   375
         Left            =   5760
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   180
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C0E1
         Caption         =   "frmYpmf090.frx":C101
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":C16F
         Keys            =   "frmYpmf090.frx":C18D
         Spin            =   "frmYpmf090.frx":C1D7
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
      Begin imNumber6Ctl.imNumber imnTax_Total 
         Height          =   375
         Left            =   11630
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   180
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C1FF
         Caption         =   "frmYpmf090.frx":C21F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":C28D
         Keys            =   "frmYpmf090.frx":C2AB
         Spin            =   "frmYpmf090.frx":C2F5
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
         Left            =   10410
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   180
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C31D
         Caption         =   "frmYpmf090.frx":C33D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":C3AB
         Keys            =   "frmYpmf090.frx":C3C9
         Spin            =   "frmYpmf090.frx":C413
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
         Left            =   7980
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   180
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1411
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C43B
         Caption         =   "frmYpmf090.frx":C45B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":C4C9
         Keys            =   "frmYpmf090.frx":C4E7
         Spin            =   "frmYpmf090.frx":C531
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
      Begin imNumber6Ctl.imNumber imnGtotal_Total 
         Height          =   375
         Left            =   13350
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   180
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C559
         Caption         =   "frmYpmf090.frx":C579
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":C5E7
         Keys            =   "frmYpmf090.frx":C605
         Spin            =   "frmYpmf090.frx":C64F
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
      Begin imNumber6Ctl.imNumber imnCharge_Total 
         Height          =   375
         Left            =   6970
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   180
         Width           =   1000
         _Version        =   65536
         _ExtentX        =   1764
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C677
         Caption         =   "frmYpmf090.frx":C697
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":C705
         Keys            =   "frmYpmf090.frx":C723
         Spin            =   "frmYpmf090.frx":C76D
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
      Begin imNumber6Ctl.imNumber imnRrate_Total 
         Height          =   375
         Left            =   8790
         TabIndex        =   235
         TabStop         =   0   'False
         Top             =   180
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1411
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C795
         Caption         =   "frmYpmf090.frx":C7B5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":C823
         Keys            =   "frmYpmf090.frx":C841
         Spin            =   "frmYpmf090.frx":C88B
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
      Begin imNumber6Ctl.imNumber imnEf_Total 
         Height          =   375
         Left            =   9600
         TabIndex        =   236
         TabStop         =   0   'False
         Top             =   180
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1411
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C8B3
         Caption         =   "frmYpmf090.frx":C8D3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":C941
         Keys            =   "frmYpmf090.frx":C95F
         Spin            =   "frmYpmf090.frx":C9A9
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
      Begin imNumber6Ctl.imNumber imnStamp_Total 
         Height          =   375
         Left            =   12540
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   180
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1411
         _ExtentY        =   661
         Calculator      =   "frmYpmf090.frx":C9D1
         Caption         =   "frmYpmf090.frx":C9F1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf090.frx":CA5F
         Keys            =   "frmYpmf090.frx":CA7D
         Spin            =   "frmYpmf090.frx":CAC7
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
End
Attribute VB_Name = "frmYpmf090"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DETAIL_MAX = 10               'ñæç◊ï\é¶çsêî
Private Const BACK_COLOR_ON = &HFF0000
Private Const BACK_COLOR_OFF = &H8000000F
Private Const BACK_COLOR_CAUTION = &HFF&
Private Const DIV_NAME1 = "ñ¢"
Private Const DIV_NAME2 = "çœ"
Private Const DIV_NAME3 = "ñ¢ï•"

Private Type Detail_Record
    Pnum As Integer
    PnumName As String
    Num As Integer
    F As Integer
    F_Total As Integer
    Total As Currency
    Charge As Currency
    Total2 As Currency
    Tax As Currency
    Keep As Currency
    Gtotal As Currency
    Div As String
    Rrate As Currency   '201107
    Ef As Currency      '201107
    Stamp As Currency   '201107
End Type
Private m_typDetail_Rec() As Detail_Record

Private Sub cboPnum_Click(Index As Integer)

    Call cboPnum_Validate(Index, False)
    
End Sub

Private Sub cboPnum_DropDown(Index As Integer)

    Call MakecboPnum(cboPnum(Index))
    
End Sub

Private Sub cboPnum_GotFocus(Index As Integer)

    cboPnum(Index).BackColor = FOCUS_STOP_COLOR
    cboPnum(Index).Tag = cboPnum(Index).Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboPnum_LostFocus(Index As Integer)
   
    cboPnum(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboPnum_Validate(Index As Integer, Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cboPnum_Validate_Err
    
    If Trim(cboPnum(Index).Text) = "" Then Exit Sub
    If cboPnum(Index).Tag = cboPnum(Index).Text Then Exit Sub
    
    lblPnum_Name(Index).Caption = ""
    
    With adoRecordset1
        'éÛïtÉfÅ[É^
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & Trim(lblOdate.Caption) & "'" & _
                 " AND Pnum = " & Trim(cboPnum(Index).Text)
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            lblPnum_Name(Index).Caption = IIf(IsNull(.Fields("Sname")), "", Trim(.Fields("Sname")))
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If Index = 0 Then
        cboPnum(1).Text = cboPnum(0).Text
        lblPnum_Name(1).Caption = lblPnum_Name(0).Caption
    End If
    
    Exit Sub

cboPnum_Validate_Err:

    Call MsgBox("ÉtÉHÅ[ÉJÉXà⁄ìÆëOÉGÉâÅ[ÅIÅI" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboPnum_Validate_Err")

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFâÊñ ÉNÉäÉAÉNÉäÉbÉNéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Sub cmdClear_Click()

    Call FieldsClear(1)

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFé¿çsÉNÉäÉbÉNéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If MsgBox("é¿çsÇµÇ‹Ç∑Ç©ÅH", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    Exit Sub
    
cmdExecute_Click_Err:

    Call MsgBox("é¿çsÉNÉäÉbÉNéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFèIóπÉNÉäÉbÉNéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdF_Search_Click(Index As Integer)

    frmView.m_intPnum = lblPnum(Index).Caption
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

    Call MsgBox("äJç√îNåéì˙Ç∆íSìñé“ÇÃïœçXÉNÉäÉbÉNéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdLogin_Click_Err")

End Sub

Private Sub cmdPrint_Click(Index As Integer)

    If MsgBox("èoïié“ê∏éZì`ï[ÇàÛç¸ÇµÇ‹Ç∑Ç©ÅH", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    Dim strArg As String
    strArg = lblOdate.Caption & "," & g_strPcode & "," & g_strPname & "," & lblPnum(Index).Caption & "," & imnNum(Index).Value
'    Call Shell(g_clsReg.Bin & "\YPMF040.exe " & strArg, vbNormalFocus)
    Call Shell(g_clsReg.Bin & "\YPMF040.exe " & strArg, vbMaximizedFocus)

End Sub

Private Sub cmdRelease_Click()

    If MsgBox("èoïié“ê∏éZÇëSâèúÇµÇ‹Ç∑Ç©ÅH", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub
    DoEvents
    If MsgBox("ñ{ìñÇ…ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub
    DoEvents
    
    If Release_Data() = False Then Exit Sub
    Call FieldsClear(1)

End Sub

Private Sub cmdSearch_Click()

    Dim intIndex1 As Integer

    For intIndex1 = 1 To DETAIL_MAX
        Call Detail_Clear(intIndex1)
    Next intIndex1
    
    Erase m_typDetail_Rec   'îzóÒèâä˙âª
    
    imnTotal_Total.Value = 0
    imnCharge_Total.Value = 0
    imnTotal2_Total.Value = 0
    imnTax_Total.Value = 0
    imnKeep_Total.Value = 0
    imnGtotal_Total.Value = 0
    
    '201107
    imnRrate_Total.Value = 0
    imnEf_Total.Value = 0
    imnStamp_Total.Value = 0
    
    Call Detail_SetData
    Call Detail_Dislplay(1)
    Call Detail_ScrollBar
    
    If UBound(m_typDetail_Rec) <= 0 Then
        Call MsgBox("ÉfÅ[É^Ç™Ç†ÇËÇ‹ÇπÇÒÅB", vbOKOnly + vbInformation, "")
    End If
    
End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFÉtÉHÅ[ÉÄÉLÅ[É_ÉEÉìéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
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

    Call MsgBox("ÉtÉHÅ[ÉÄÉLÅ[É_ÉEÉìéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")

End Sub

'ñ⁄Å@ìIÅ@Å@ÅF
'èÅ@åèÅ@Å@ÅFÉtÉHÅ[ÉÄÉçÅ[Éhéû
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "èoïié“ê∏éZèÛãµ"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    Call FieldsClear(0)
    Timer1.Enabled = True
    
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
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set g_clsAdoSQL = Nothing
    Set g_clsReg = Nothing
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
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Sub FieldsClear(intKubun As Integer)

    Dim intIndex1 As Integer

    On Error GoTo FieldsClear_Err
    
    cboPnum(0).Text = ""
    cboPnum(1).Text = ""
    lblPnum_Name(0).Caption = ""
    lblPnum_Name(1).Caption = ""
    
    For intIndex1 = 1 To DETAIL_MAX
        Call Detail_Clear(intIndex1)
    Next intIndex1
    
    Erase m_typDetail_Rec   'îzóÒèâä˙âª
    
    imnTotal_Total.Value = 0
    imnCharge_Total.Value = 0
    imnTotal2_Total.Value = 0
    imnTax_Total.Value = 0
    imnKeep_Total.Value = 0
    imnGtotal_Total.Value = 0
    
    '201107
    imnRrate_Total.Value = 0
    imnEf_Total.Value = 0
    imnStamp_Total.Value = 0
    
    If intKubun = 1 Then
        Call Detail_SetData
        Call Detail_Dislplay(1)
        Call Detail_ScrollBar
        
        If UBound(m_typDetail_Rec) <= 0 Then
            Call MsgBox("ÉfÅ[É^Ç™Ç†ÇËÇ‹ÇπÇÒÅB", vbOKOnly + vbInformation, "")
        End If
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("âÊñ ÉNÉäÉAÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

'ñ⁄Å@ìIÅ@Å@ÅFì¸óÕÉ`ÉFÉbÉN
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "äJç√îNåéì˙Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
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

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cboPnum(0).SetFocus

End Sub

'ñ⁄Å@ìIÅ@Å@ÅFñæç◊ÉèÅ[ÉNÇ÷ÉfÅ[É^ÉZÉbÉg
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Function Detail_SetData() As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intRecordCount As Integer
    Dim blnFlg As Boolean
    Dim typDetail_Sort() As Detail_Record
    Dim strBuff As String

    On Error GoTo Detail_SetData_Err
    
    Detail_SetData = False
    
    Screen.MousePointer = vbHourglass
    
    'èâä˙âª
    ReDim typDetail_Sort(0): ReDim m_typDetail_Rec(0)
    
    'éÛïtÉfÅ[É^
    If cboPnum(0).Text <> "" And cboPnum(1).Text <> "" Then
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum BETWEEN " & cboPnum(0).Text & " AND " & cboPnum(1).Text & _
                 " ORDER BY Odate,Pnum"
    Else
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " ORDER BY Odate,Pnum"
    End If
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        intRecordCount = adoRecordset1.RecordCount
        ReDim typDetail_Sort(intRecordCount)
    End If
    
    For intIndex1 = 1 To intRecordCount
        'èoïié“ê∏éZÉfÅ[É^
        strSQL = "SELECT * FROM DT040" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & adoRecordset1.Fields("Pnum") & _
                 " ORDER BY Odate,Pnum,Num"
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoRecordset2.EOF = False Then
            Do While Not adoRecordset2.EOF
                typDetail_Sort(intIndex1).Pnum = IIf(IsNull(adoRecordset1.Fields("Pnum")), 0, adoRecordset1.Fields("Pnum"))
                typDetail_Sort(intIndex1).PnumName = IIf(IsNull(adoRecordset1.Fields("Sname")), "", adoRecordset1.Fields("Sname"))
                typDetail_Sort(intIndex1).Num = adoRecordset2.Fields("Num")
                typDetail_Sort(intIndex1).Total = typDetail_Sort(intIndex1).Total + CCur(adoRecordset2.Fields("Total")) + CCur(adoRecordset2.Fields("Ototal"))
                typDetail_Sort(intIndex1).Charge = typDetail_Sort(intIndex1).Charge + CCur(adoRecordset2.Fields("Charge"))
                '202308 ÉCÉìÉ{ÉCÉXëŒâû typDetail_Sort(intIndex1).Total2 = typDetail_Sort(intIndex1).Total - typDetail_Sort(intIndex1).Charge
                typDetail_Sort(intIndex1).Tax = typDetail_Sort(intIndex1).Tax + CCur(adoRecordset2.Fields("Tax"))
                typDetail_Sort(intIndex1).Keep = typDetail_Sort(intIndex1).Keep + CCur(adoRecordset2.Fields("Keep"))
                typDetail_Sort(intIndex1).Gtotal = typDetail_Sort(intIndex1).Gtotal + CCur(adoRecordset2.Fields("Gtotal"))
                If Not IsNull(adoRecordset2.Fields("Pdiv")) Then
                    If adoRecordset2.Fields("Pdiv") = SHIHARAI_ON Then
                        typDetail_Sort(intIndex1).Div = DIV_NAME2
                    Else
                        typDetail_Sort(intIndex1).Div = DIV_NAME3
                    End If
                Else
                    typDetail_Sort(intIndex1).Div = DIV_NAME2
                End If
                
                '201107
                typDetail_Sort(intIndex1).Rrate = typDetail_Sort(intIndex1).Rrate + CCur(adoRecordset2.Fields("Rrate"))
                typDetail_Sort(intIndex1).Ef = typDetail_Sort(intIndex1).Ef + CCur(adoRecordset2.Fields("Ef"))
                typDetail_Sort(intIndex1).Stamp = typDetail_Sort(intIndex1).Stamp + CCur(adoRecordset2.Fields("Stamp"))
                '202308 ÉCÉìÉ{ÉCÉXëŒâû
                typDetail_Sort(intIndex1).Total2 = typDetail_Sort(intIndex1).Total - typDetail_Sort(intIndex1).Charge - typDetail_Sort(intIndex1).Keep - typDetail_Sort(intIndex1).Rrate - typDetail_Sort(intIndex1).Ef

                adoRecordset2.MoveNext
            Loop
        Else
            typDetail_Sort(intIndex1).Pnum = IIf(IsNull(adoRecordset1.Fields("Pnum")), 0, adoRecordset1.Fields("Pnum"))
            typDetail_Sort(intIndex1).PnumName = IIf(IsNull(adoRecordset1.Fields("Sname")), "", adoRecordset1.Fields("Sname"))
            typDetail_Sort(intIndex1).Num = 0
            typDetail_Sort(intIndex1).Total = 0
            typDetail_Sort(intIndex1).Charge = 0
            typDetail_Sort(intIndex1).Total2 = 0
            typDetail_Sort(intIndex1).Tax = 0
            typDetail_Sort(intIndex1).Keep = 0
            typDetail_Sort(intIndex1).Gtotal = 0
            typDetail_Sort(intIndex1).Div = DIV_NAME1
            
            '201107
            typDetail_Sort(intIndex1).Rrate = 0
            typDetail_Sort(intIndex1).Ef = 0
            typDetail_Sort(intIndex1).Stamp = 0

        End If
        adoRecordset2.Close
        
        adoRecordset1.MoveNext
    Next intIndex1
    adoRecordset1.Close
    
    'èoïié“ÉRÅ[ÉhÇ≈É\Å[Ég
    Call Detail_Sort(typDetail_Sort, m_typDetail_Rec)
    
    'FñáêîÇÃéÊìæ
    Call Get_F
    
    'è¨åvÇ»Ç«ÇÃåvéZ
    Call Calc_Total
    
    Screen.MousePointer = vbDefault
    
    Detail_SetData = True
    
    Exit Function

Detail_SetData_Err:

    Detail_SetData = False
    Screen.MousePointer = vbDefault
    Call MsgBox("ÉtÉBÅ[ÉãÉhÉZÉbÉgÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_SetData_Err")

End Function

'ñ⁄Å@ìIÅ@Å@ÅFçáåvÇÃåvéZ
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Sub Calc_Total()

    Dim curBuff(9) As Currency
    Dim intIndex1 As Integer

    On Error GoTo Calc_Total_Err
    
    curBuff(1) = 0: curBuff(2) = 0: curBuff(3) = 0: curBuff(4) = 0: curBuff(5) = 0: curBuff(6) = 0: curBuff(7) = 0: curBuff(8) = 0: curBuff(9) = 0
    
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
         curBuff(1) = curBuff(1) + m_typDetail_Rec(intIndex1).Total
         curBuff(2) = curBuff(2) + m_typDetail_Rec(intIndex1).Charge
         curBuff(3) = curBuff(3) + m_typDetail_Rec(intIndex1).Total2
         curBuff(4) = curBuff(4) + m_typDetail_Rec(intIndex1).Tax
         curBuff(5) = curBuff(5) + m_typDetail_Rec(intIndex1).Keep
         curBuff(6) = curBuff(6) + m_typDetail_Rec(intIndex1).Gtotal
         '201107
         curBuff(7) = curBuff(7) + m_typDetail_Rec(intIndex1).Rrate
         curBuff(8) = curBuff(8) + m_typDetail_Rec(intIndex1).Ef
         curBuff(9) = curBuff(9) + m_typDetail_Rec(intIndex1).Stamp
    Next intIndex1
    
    imnTotal_Total.Value = curBuff(1)
    imnCharge_Total.Value = curBuff(2)
    imnTotal2_Total.Value = curBuff(3)
    imnTax_Total.Value = curBuff(4)
    imnKeep_Total.Value = curBuff(5)
    imnGtotal_Total.Value = curBuff(6)
    
    '201107
    imnRrate_Total.Value = curBuff(7)
    imnEf_Total.Value = curBuff(8)
    imnStamp_Total.Value = curBuff(9)
    
    Exit Sub
    
Calc_Total_Err:

    Call MsgBox("çáåvÇÃåvéZÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Total_Err")

End Sub

'ñ⁄Å@ìIÅ@Å@ÅFñæç◊ÇÃï\é¶
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF(intStartLineÅFï\é¶äJénçsî‘çÜ)
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Function Detail_Dislplay(intStartLine As Integer) As Boolean

    Dim intIndex1 As Integer
    Dim intPostion As Integer

    On Error GoTo Detail_Dislplay_Err
    
    Screen.MousePointer = vbHourglass
    
    Detail_Dislplay = False
    
    intPostion = intStartLine
    For intIndex1 = 1 To DETAIL_MAX
        'ñæç◊ÇÃÇPçsÉNÉäÉA
        Call Detail_Clear(intIndex1)
        
        If intPostion <= UBound(m_typDetail_Rec) Then
            fraDetail(intIndex1).Visible = True
        
            lblPnum(intIndex1).Caption = m_typDetail_Rec(intPostion).Pnum
            lblPnumName(intIndex1).Caption = m_typDetail_Rec(intPostion).PnumName
            imnNum(intIndex1).Value = m_typDetail_Rec(intPostion).Num
            imnF(intIndex1).Value = m_typDetail_Rec(intPostion).F
            imnFtotal(intIndex1).Value = m_typDetail_Rec(intPostion).F_Total
            imnTotal(intIndex1).Value = m_typDetail_Rec(intPostion).Total
            imnCharge(intIndex1).Value = m_typDetail_Rec(intPostion).Charge
            imnTotal2(intIndex1).Value = m_typDetail_Rec(intPostion).Total2
            imnTax(intIndex1).Value = m_typDetail_Rec(intPostion).Tax
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
            
            '201107
            imnRrate(intIndex1).Value = m_typDetail_Rec(intPostion).Rrate
            imnEf(intIndex1).Value = m_typDetail_Rec(intPostion).Ef
            imnStamp(intIndex1).Value = m_typDetail_Rec(intPostion).Stamp

        End If
        intPostion = intPostion + 1
    Next intIndex1
    
    Screen.MousePointer = vbDefault
    
    Detail_Dislplay = True
    
    Exit Function
    
Detail_Dislplay_Err:

    Screen.MousePointer = vbDefault
    Detail_Dislplay = False
    Call MsgBox("ñæç◊ÇÃï\é¶ÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Dislplay_Err")

End Function

'ñ⁄Å@ìIÅ@Å@ÅFñæç◊ÇÃÇPçsÉNÉäÉA
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF(intClearLine:ê‚ëŒà íu)
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Function Detail_Clear(intClearLine As Integer) As Boolean

    On Error GoTo Detail_Clear_Err
    
    lblPnum(intClearLine).Caption = ""
    lblPnumName(intClearLine).Caption = ""
    imnNum(intClearLine).Value = 0
    imnF(intClearLine).Value = 0
    imnFtotal(intClearLine).Value = 0
    imnTotal(intClearLine).Value = 0
    imnCharge(intClearLine).Value = 0
    imnTotal2(intClearLine).Value = 0
    imnTax(intClearLine).Value = 0
    imnKeep(intClearLine).Value = 0
    imnGtotal(intClearLine).Value = 0
    lblDiv(intClearLine).Caption = ""
    fraDetail(intClearLine).BackColor = BACK_COLOR_OFF
    fraDetail(intClearLine).Visible = False
    
    '201107
    imnRrate(intClearLine).Value = 0
    imnEf(intClearLine).Value = 0
    imnStamp(intClearLine).Value = 0
    
    Detail_Clear = True
    
    Exit Function
    
Detail_Clear_Err:

    Detail_Clear = False
    Call MsgBox("ñæç◊ÇÃÇPçsÉNÉäÉAÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Clear_Err")

End Function

'ñ⁄Å@ìIÅ@Å@ÅFÉXÉNÉçÅ[ÉãÉoÅ[ÇÃêßå‰
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
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
    Call MsgBox("ÉXÉNÉçÅ[ÉãÉoÅ[ÇÃêßå‰ÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_ScrollBar_Err")

End Function

Private Sub lblDiv_DblClick(Index As Integer)

    On Error GoTo lblDiv_DblClick_Err

    If Trim(lblDiv(Index).Caption) <> DIV_NAME2 Then Exit Sub
    
    If MsgBox("éÛïtî‘çÜÅF" & lblPnum(Index).Caption & "ÇÃê∏éZÇâèúÇµÇ‹Ç∑Ç©ÅH", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub
    
    If IsNumeric(lblPnum(Index).Caption) Then
        If Release_Data(CInt(lblPnum(Index).Caption)) = False Then Exit Sub
    End If
    Call FieldsClear(1)

    Exit Sub

lblDiv_DblClick_Err:

    Call MsgBox("ê∏éZâèúÉGÉâÅ[ÅIÅI" _
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

'ñ⁄Å@ìIÅ@Å@ÅFñæç◊ÇÃÉ\Å[Ég
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Function Detail_Sort(ByRef Before() As Detail_Record, ByRef After() As Detail_Record) As Boolean

    Dim intIndex1 As Integer
    Dim intPostion As Integer
    Dim work(1) As Detail_Record     'ÉèÅ[ÉN

    On Error GoTo Detail_Sort_Err
    
    If UBound(Before) <= 0 Then Exit Function
    
    'ÉoÉuÉãÉ\Å[Ég
    For intIndex1 = UBound(Before) To 1 Step -1
        For intPostion = 1 To intIndex1 - 1
            If CInt(Before(intPostion).Pnum) >= CInt(Before(intPostion + 1).Pnum) And CInt(Before(intPostion).Num) >= CInt(Before(intPostion + 1).Num) Then
                work(1) = Before(intPostion)
                Before(intPostion) = Before(intPostion + 1)
                Before(intPostion + 1) = work(1)
            End If
        Next intPostion
    Next intIndex1
    
    'îzóÒÉRÉsÅ[
    ReDim After(UBound(Before))
    After = Before
    
    Detail_Sort = True
    
    Exit Function
    
Detail_Sort_Err:

    Detail_Sort = False
    Call MsgBox("ñæç◊ÇÃÉ\Å[ÉgÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Sort_Err")

End Function

'ñ⁄Å@ìIÅ@Å@ÅFèoóÕçœÇ›ÇÃâèú
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Function Release_Data(Optional intPnum As Variant) As Boolean

    Dim strSQL As String

    On Error GoTo Release_Data_Err
    
    Screen.MousePointer = vbHourglass
    
    Release_Data = False
    
    With g_clsAdoSQL.Connection
        .BeginTrans
        
        If IsMissing(intPnum) = False Then
            '********** éwíËèoïié“ÇÃâèú **********
        
            'ã£îÑñæç◊ÉfÅ[É^
            strSQL = "UPDATE DT021" & _
                     " SET Sdiv = " & EXHIBITION_REPORT_OFF & "," & _
                     " Snum = 0" & _
                     " WHERE LEFT(Ocode,8) = '" & Global_Get_NumericDay(lblOdate.Caption) & "'" & _
                     " AND Pnum = " & intPnum
            .Execute strSQL
        
            'éÛïtñæç◊ÉfÅ[É^
            strSQL = "UPDATE DT011" & _
                     " SET Sdiv = " & EXHIBITION_REPORT_OFF & "," & _
                     " Snum = 0" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Pnum = " & intPnum
            .Execute strSQL
        
            'èoïié“ê∏éZÉfÅ[É^ÇÃçÌèú
            strSQL = "DELETE FROM DT040" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                     " AND Pnum = " & intPnum
            .Execute strSQL
        Else
            '********** ëSåèâèú **********
        
            'ã£îÑñæç◊ÉfÅ[É^
            strSQL = "UPDATE DT021" & _
                     " SET Sdiv = " & EXHIBITION_REPORT_OFF & "," & _
                     " Snum = 0" & _
                     " WHERE LEFT(Ocode,8) = '" & Global_Get_NumericDay(lblOdate.Caption) & "'"
            .Execute strSQL
        
            'éÛïtñæç◊ÉfÅ[É^
            strSQL = "UPDATE DT011" & _
                     " SET Sdiv = " & EXHIBITION_REPORT_OFF & "," & _
                     " Snum = 0" & _
                     " WHERE Odate = '" & lblOdate.Caption & "'"
            .Execute strSQL
        
            'èoïié“ê∏éZÉfÅ[É^ÇÃçÌèú
            strSQL = "DELETE FROM DT040" & _
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
    Call MsgBox("âèúÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Release_Data_Err")

End Function

'ñ⁄Å@ìIÅ@Å@ÅFFñáêîÇÃéÊìæ
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇWÅ^ÇPÇQ
'çXêVóöóÅ@ÅF
'
Private Sub Get_F()

    Dim intIndex1 As Integer
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strOdate As String

    On Error GoTo Get_F_Err
    
    strOdate = CStr(Global_Get_NumericDay(lblOdate.Caption))
    
    For intIndex1 = 1 To UBound(m_typDetail_Rec)
        'ã£îÑñæç◊ÉfÅ[É^
        strSQL = "SELECT * FROM DT021" & _
                 " WHERE LEFT(Ocode,8) = '" & strOdate & "'" & _
                 " AND Pnum = " & m_typDetail_Rec(intIndex1).Pnum
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoRecordset1.EOF = False Then
            m_typDetail_Rec(intIndex1).F = adoRecordset1.RecordCount
        Else
            m_typDetail_Rec(intIndex1).F = 0
        End If
        adoRecordset1.Close
        
'        'éÛïtñæç◊ÉfÅ[É^(íçï∂ï™Çâ¡éZ)
'        strSQL = "SELECT * FROM DT011" & _
'                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
'                 " AND Pnum = " & m_typDetail_Rec(intIndex1).Pnum & _
'                 " AND Price <> 0"
'        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
'        If adoRecordset1.EOF = False Then
'            m_typDetail_Rec(intIndex1).F = m_typDetail_Rec(intIndex1).F + CCur(adoRecordset1.RecordCount)
'        End If
'        adoRecordset1.Close
        
        'éÛïtñæç◊ÉfÅ[É^
        strSQL = "SELECT * FROM DT011" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & m_typDetail_Rec(intIndex1).Pnum
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoRecordset1.EOF = False Then
            m_typDetail_Rec(intIndex1).F_Total = adoRecordset1.RecordCount
        Else
            m_typDetail_Rec(intIndex1).F_Total = 0
        End If
        adoRecordset1.Close
    Next intIndex1
    
    Exit Sub
    
Get_F_Err:

    Call MsgBox("FñáêîÇÃéÊìæÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Get_F_Err")

End Sub

'ñ⁄Å@ìIÅ@Å@ÅFÉRÉìÉ{É{ÉbÉNÉXÇÃçÏê¨
'èÅ@åèÅ@Å@ÅF
'åãÅ@â Å@Å@ÅF
'à¯Å@êîÅ@Å@ÅF
'ñﬂÇËílÅ@Å@ÅF
'çÏê¨é“Å@Å@ÅFäîéÆâÔé– ÉRÉÄÅEÉGÉìÉWÉjÉAÉäÉìÉOÅ@à≠î¸
'çÏê¨îNåéì˙ÅFÇQÇOÇOÇQÅ^ÇOÇVÅ^ÇQÇT
'çXêVóöóÅ@ÅF
'
Private Sub MakecboPnum(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboPnum_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & Trim(lblOdate.Caption) & "'"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Ctrl.AddItem .Fields("Pnum") & ";" & IIf(IsNull(.Fields("Sname")), "", .Fields("Sname"))
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboPnum_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("ÉRÉìÉ{É{ÉbÉNÉXçÏê¨ÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboPnum_Err")

End Sub


