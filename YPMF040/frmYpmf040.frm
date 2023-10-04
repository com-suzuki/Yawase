VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmYpmf040 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   4320
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf040.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   12150
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdLogin 
         Caption         =   "開催年月日と担当者の変更"
         Height          =   375
         Left            =   6960
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9900
         TabIndex        =   18
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
         Contents        =   "frmYpmf040.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   15
      Top             =   3540
      Width           =   12015
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   9
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
         Picture         =   "frmYpmf040.frx":0D13
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10200
         TabIndex        =   11
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
         Picture         =   "frmYpmf040.frx":0D2F
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   8460
         TabIndex        =   10
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
         Picture         =   "frmYpmf040.frx":0E89
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   2895
      Left            =   60
      TabIndex        =   14
      Top             =   660
      Width           =   12015
      Begin VB.Frame fraReprint 
         BorderStyle     =   0  'なし
         Height          =   495
         Left            =   2460
         TabIndex        =   30
         Top             =   2280
         Width           =   3435
         Begin imNumber6Ctl.imNumber imnRePrintNum 
            Height          =   435
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   767
            Calculator      =   "frmYpmf040.frx":0F9B
            Caption         =   "frmYpmf040.frx":0FBB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf040.frx":1029
            Keys            =   "frmYpmf040.frx":1047
            Spin            =   "frmYpmf040.frx":1091
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2011496453
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnRePrintNum 
            Height          =   435
            Index           =   1
            Left            =   1860
            TabIndex        =   8
            Top             =   0
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   767
            Calculator      =   "frmYpmf040.frx":10B9
            Caption         =   "frmYpmf040.frx":10D9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf040.frx":1147
            Keys            =   "frmYpmf040.frx":1165
            Spin            =   "frmYpmf040.frx":11AF
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2011496453
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin VB.Label lblRePrintNum 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "回目 ～"
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
            Index           =   0
            Left            =   660
            TabIndex        =   32
            Top             =   60
            Width           =   1095
         End
         Begin VB.Label lblRePrintNum 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "回目"
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
            Index           =   1
            Left            =   2520
            TabIndex        =   31
            Top             =   60
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'なし
         Height          =   615
         Left            =   1620
         TabIndex        =   29
         Top             =   1500
         Width           =   4335
         Begin VB.OptionButton optFdiv 
            Caption         =   "銀行振込"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   2880
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   5
            Top             =   120
            Width           =   1395
         End
         Begin VB.OptionButton optFdiv 
            Caption         =   "現　金"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   0
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   3
            Top             =   120
            Width           =   1395
         End
         Begin VB.OptionButton optFdiv 
            Caption         =   "小切手"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   1440
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   4
            Top             =   120
            Width           =   1395
         End
      End
      Begin VB.CheckBox chkRePrint 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Width           =   555
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1620
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "支払種別"
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
         Contents        =   "frmYpmf040.frx":11D7
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "受付番号"
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
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   2220
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "再発行"
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
      Begin CSComboLib.CSComboBox cboPnum 
         Height          =   405
         Index           =   1
         Left            =   1620
         TabIndex        =   2
         Top             =   1020
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
         Contents        =   "frmYpmf040.frx":11F0
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
         Left            =   1620
         TabIndex        =   28
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblPnum_Name 
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
         Left            =   2700
         TabIndex        =   27
         Top             =   1020
         Width           =   9195
      End
      Begin VB.Label lblPnum_Name 
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
         TabIndex        =   26
         Top             =   180
         Width           =   9195
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   12240
      TabIndex        =   0
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf040.frx":1209
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf040.frx":1277
      Key             =   "frmYpmf040.frx":1295
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
      Left            =   12240
      TabIndex        =   12
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf040.frx":12D9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf040.frx":1347
      Key             =   "frmYpmf040.frx":1365
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
      TabIndex        =   13
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf040.frx":13A9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf040.frx":1417
      Key             =   "frmYpmf040.frx":1435
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
Attribute VB_Name = "frmYpmf040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    If IsNumeric(cboPnum(Index).Text) = False Then
        cboPnum(Index).Text = ""
        lblPnum_Name(Index).Caption = ""
        Exit Sub
    End If
    If cboPnum(Index).Tag = cboPnum(Index).Text Then Exit Sub
    
    lblPnum_Name(Index).Caption = ""
    
    With adoRecordset1
        '受付データ
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

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboPnum_Validate_Err")

End Sub

Private Sub chkRePrint_Click()

'    If chkRePrint.Value = 1 Then
'        fraReprint.Visible = True
'        imnRePrintNum(0).SetFocus
'    Else
'        fraReprint.Visible = False
'    End If
    
End Sub

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２２
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    cboPnum(0).SetFocus

End Sub

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２２
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If MsgBox("実行しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    '入力チェック
    If DoValidationChecks() = False Then Exit Sub
    '印刷用ワーク作成
    If MakePrintWork() = False Then Exit Sub
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
'作成年月日：２００２／０７／２２
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
'作成年月日：２００２／０７／２２
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
'作成年月日：２００２／０７／２２
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "出品者精算"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    Call FieldsClear(0)
    
    If Command() <> "" Then
        '入力チェック
        If DoValidationChecks() = False Then End
        '印刷用ワーク作成
        If MakePrintWork() = False Then End
        '印刷
        If ActiveReportPrint(0) = False Then End
        End
    End If
    
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
'作成年月日：２００２／０７／２２
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
'作成年月日：２００２／０７／２２
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        '2005/09/14 修正
        If Trim$(g_strPnum) = "" Then
            cboPnum(0).Text = ""
            cboPnum(0).Tag = ""
            cboPnum(1).Text = ""
            cboPnum(1).Tag = ""
        Else
            cboPnum(0).Text = g_strPnum
            cboPnum(0).Tag = g_strPnum
            cboPnum(1).Text = g_strPnum
            cboPnum(1).Tag = g_strPnum
        End If
    
        lblPnum_Name(0).Caption = ""
        lblPnum_Name(1).Caption = ""
        optFdiv(0).Value = True
        
        If Trim$(g_strRePrintNum) = "0" Then
            chkRePrint.Value = 0
            fraReprint.Visible = False
            imnRePrintNum(0).Value = 0
            imnRePrintNum(1).Value = 99
        Else
            chkRePrint.Value = 1
            fraReprint.Visible = True
            imnRePrintNum(0).Value = g_strRePrintNum
            imnRePrintNum(1).Value = g_strRePrintNum
        End If
        
        fraReprint.Visible = False
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
'作成年月日：２００２／０７／２２
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "開催年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(cboPnum(0).Text) = "" Then
        cboPnum(0).SetFocus
        strErrMsg = "受付番号を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(cboPnum(1).Text) = "" Then
        cboPnum(1).SetFocus
        strErrMsg = "受付番号を入力してください。"
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
'作成年月日：２００２／０７／２２
'更新履歴　：
'
Private Function MakePrintWork_Old() As Boolean

    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim adoDT010 As New ADODB.Recordset
    Dim adoDT011 As New ADODB.Recordset
    Dim adoDT021 As New ADODB.Recordset
    Dim adoDT030 As New ADODB.Recordset
    Dim adoDT040 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim wkRecordsetTemp As New ADODB.Recordset
    Dim strBuff1 As String

    Dim intLine As Integer                  '行番号
    Dim lngCount As Long                    'レコード件数
    Dim curSkeep As Currency                '出品者維持管理費(標準)
    Dim curSkeepCurrent As Currency         '出品者維持管理費(今回)
    Dim curSrate As Currency                '出品者手数料率
    Dim intSfraction As Integer             '出品者端数処理
    Dim intNum As Integer                   '回数
    Dim curTaxRate As Currency              '消費税率
    Dim intR As Integer                     '支払種別
    Dim curSRounding As Currency            '出品者丸め単位
    Dim curTotalPrice As Currency           '金額(競売金額＋注文金額)
    Dim varKey1 As Variant
    Dim curBuff As Currency
    Dim blnFlg As Boolean

    On Error GoTo MakePrintWork_Old_Err
    
    MakePrintWork_Old = False
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
'********** 初期処理 **********
    
    '初期化
    lngCount = 0        'レコード件数
    curSkeep = 0        '出品者維持管理費
    intSfraction = 0    '出品者端数処理
    curSrate = 0        '出品者手数料率
    curTaxRate = 0      '消費税率
    intR = 0            '支払種別
    curSRounding = 0    '出品者丸め単位
    
    If optFdiv(0).Value = True Then
        intR = PAYMENT_DIV_CASH
    ElseIf optFdiv(1).Value = True Then
        intR = PAYMENT_DIV_CHECK
    ElseIf optFdiv(2).Value = True Then
        intR = PAYMENT_DIV_TRANSFER
    End If
    
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
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, Trim(lblOdate.Caption))
    
'********** ワーク **********
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF040"
    g_clsAdoAccess.Connection.Execute strSQL

    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF040"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '受付データオープン
    strSQL = "{call sp_YPMF0401;1('" & Trim(lblOdate.Caption) & "'," & _
              cboPnum(0).Text & "," & cboPnum(1).Text & ")}"
    adoDT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT010.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = adoDT010.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    
    Do While Not adoDT010.EOF
        '受付明細データオープン
        If chkRePrint.Value = 0 Then
            strSQL = "{call sp_YPMF0402;1('" & lblOdate.Caption & "'," & adoDT010.Fields("Pnum") & ")}"
        Else
            strSQL = "{call sp_YPMF0402;2('" & lblOdate.Caption & "'," & adoDT010.Fields("Pnum") & "," & _
                     imnRePrintNum(0).Text & "," & imnRePrintNum(1).Text & ")}"
        End If
        adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        Do While Not adoDT011.EOF
            If IsNull(adoDT011.Fields("Price")) = False And adoDT011.Fields("Price") <> 0 Then
                '********** 注文分 **********
                wkRecordset.AddNew
                If chkRePrint.Value = 0 Then
                    wkRecordset.Fields("Key1") = Format(adoDT010.Fields("Pnum"), "0000") & "01"
                    wkRecordset.Fields("Key2") = Format(adoDT010.Fields("Pnum"), "0000") & "01"
                    wkRecordset.Fields("Num") = 1
                    wkRecordset.Fields("Div") = "B"
                Else
                    wkRecordset.Fields("Key1") = Format(adoDT010.Fields("Pnum"), "0000") & Format(adoDT011.Fields("Snum"), "00")
                    wkRecordset.Fields("Key2") = Format(adoDT010.Fields("Pnum"), "0000") & Format(adoDT011.Fields("Snum"), "00")
                    wkRecordset.Fields("Num") = adoDT011.Fields("Snum")
                    wkRecordset.Fields("Div") = "B"
                End If
                wkRecordset.Fields("Odate") = Trim(lblOdate.Caption)
                wkRecordset.Fields("Pnum") = adoDT010.Fields("Pnum")
                wkRecordset.Fields("Scode") = adoDT010.Fields("Scode")
                wkRecordset.Fields("Sname") = adoDT010.Fields("Sname")
                wkRecordset.Fields("Line") = adoDT011.Fields("Line")
                wkRecordset.Fields("Icode") = adoDT011.Fields("Icode")
                wkRecordset.Fields("Iname") = adoDT011.Fields("Iname")
                wkRecordset.Fields("Qty") = adoDT011.Fields("Qty")
                wkRecordset.Fields("Price1") = adoDT011.Fields("Price")
                wkRecordset.Fields("Bcode") = adoDT011.Fields("Bcode")
                wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, adoDT011.Fields("Bcode"), Trim(lblOdate.Caption), strBuff1)
'                wkRecordset.Fields("Total") = 0
'                wkRecordset.Fields("Charge") = 0
'                wkRecordset.Fields("Tax") = 0
'                wkRecordset.Fields("Keep") = curSkeep
'                wkRecordset.Fields("GTotal") = 0
                wkRecordset.Fields("Itime") = Format(Now(), "hh時mm分")
                wkRecordset.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
                wkRecordset.Fields("Pname") = Trim(lblPname.Caption)
                wkRecordset.Fields("Idiv") = 0
                wkRecordset.Fields("Ocode") = Format(adoDT010.Fields("Pnum"), "0000") & "*"
                wkRecordset.Update
                
                If chkRePrint.Value = 0 Then
                    '出品者精算データから精算回数を取得
                    strSQL = "SELECT * FROM DT040" & _
                             " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                             " AND Pnum = " & wkRecordset.Fields("Pnum") & _
                             " ORDER BY Odate,Pnum,Num DESC"
                    adoDT040.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                    intNum = 1
                    If adoDT040.EOF = False Then
                        intNum = CInt(adoDT040.Fields("Num")) + 1
                    End If
                    adoDT040.Close
                    
                    '受付データ更新
                    adoDT011.Fields("Sdiv") = EXHIBITION_REPORT_ON
                    adoDT011.Fields("Snum") = intNum
                    adoDT011.Update
                End If
            End If
            adoDT011.MoveNext
        Loop
        adoDT011.Close
        
        '競売明細データオープン
        If chkRePrint.Value = 0 Then
            strSQL = "{call sp_YPMF0403;1('" & Global_Get_NumericDay(lblOdate.Caption) & "'," & adoDT010.Fields("Pnum") & ")}"
        Else
            strSQL = "{call sp_YPMF0403;2('" & Global_Get_NumericDay(lblOdate.Caption) & "'," & adoDT010.Fields("Pnum") & "," & _
                     imnRePrintNum(0).Text & "," & imnRePrintNum(1).Text & ")}"
        End If
        adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        Do While Not adoDT021.EOF
            wkRecordset.AddNew
            If chkRePrint.Value = 0 Then
                wkRecordset.Fields("Key1") = Format(adoDT010.Fields("Pnum"), "0000") & "01"
                wkRecordset.Fields("Key2") = Format(adoDT010.Fields("Pnum"), "0000") & "01"
                wkRecordset.Fields("Num") = 1
                wkRecordset.Fields("Div") = "A"
            Else
                wkRecordset.Fields("Key1") = Format(adoDT010.Fields("Pnum"), "0000") & Format(adoDT021.Fields("Snum"), "00")
                wkRecordset.Fields("Key2") = Format(adoDT010.Fields("Pnum"), "0000") & Format(adoDT021.Fields("Snum"), "00")
                wkRecordset.Fields("Num") = adoDT021.Fields("Snum")
                wkRecordset.Fields("Div") = "A"
            End If
            wkRecordset.Fields("Odate") = Trim(lblOdate.Caption)
            wkRecordset.Fields("Pnum") = adoDT010.Fields("Pnum")
            wkRecordset.Fields("Scode") = adoDT010.Fields("Scode")
            wkRecordset.Fields("Sname") = adoDT010.Fields("Sname")
            wkRecordset.Fields("Line") = adoDT021.Fields("PnumLine")
            wkRecordset.Fields("Icode") = adoDT021.Fields("Icode")
            wkRecordset.Fields("Iname") = adoDT021.Fields("Iname")
            wkRecordset.Fields("Qty") = adoDT021.Fields("Qty")
            wkRecordset.Fields("Price1") = adoDT021.Fields("Price")
            wkRecordset.Fields("Bcode") = adoDT021.Fields("Bcode")
            If Not IsNull(wkRecordset.Fields("Bcode")) Then
                wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, adoDT021.Fields("Bcode"), Trim(lblOdate.Caption), strBuff1)
            Else
                wkRecordset.Fields("Bname") = ""
            End If
'            wkRecordset.Fields("Total") = 0
'            wkRecordset.Fields("Charge") = 0
'            wkRecordset.Fields("Tax") = 0
'            wkRecordset.Fields("Keep") = curSkeep
'            wkRecordset.Fields("GTotal") = 0
            wkRecordset.Fields("Itime") = Format(Now(), "hh時mm分")
            wkRecordset.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
            wkRecordset.Fields("Pname") = Trim(lblPname.Caption)
            If IsNull(adoDT021.Fields("Sline")) = False Then
                wkRecordset.Fields("Idiv") = adoDT021.Fields("Sline")
            Else
                wkRecordset.Fields("Idiv") = 0
            End If
            If IsNull(adoDT021.Fields("Idiv")) = False Then
                wkRecordset.Fields("Result") = adoDT021.Fields("Idiv")
            Else
                wkRecordset.Fields("Result") = 0
            End If
            wkRecordset.Fields("Ocode") = right$(adoDT021.Fields("Ocode"), 4)
        
            wkRecordset.Update
        
            If chkRePrint.Value = 0 Then
                '出品者精算データから精算回数を取得
                strSQL = "SELECT * FROM DT040" & _
                         " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                         " AND Pnum = " & wkRecordset.Fields("Pnum") & _
                         " ORDER BY Odate,Pnum,Num DESC"
                adoDT040.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                intNum = 1
                If adoDT040.EOF = False Then
                    intNum = CInt(adoDT040.Fields("Num")) + 1
                End If
                adoDT040.Close
            
                '競売明細データ更新
                adoDT021.Fields("Sdiv") = EXHIBITION_REPORT_ON
                adoDT021.Fields("Snum") = intNum
                adoDT021.Update
            End If
            adoDT021.MoveNext
        Loop
        adoDT021.Close
        
        adoDT010.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Old_Cancel:
    Loop
    adoDT010.Close
    
'********** 出品者精算データ作成 **********

    wkRecordset.Close
    lngCount = 0
    
    'ワークオープン
    If chkRePrint.Value = 0 Then
        strSQL = "SELECT WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname, Sum(WK_YPMF040.Price1) AS Price1_Total" & _
                 " FROM WK_YPMF040 " & _
                 " GROUP BY WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname" & _
                 " ORDER BY WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname"
    Else
        strSQL = "SELECT WK_YPMF040.Key1, WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname, Sum(WK_YPMF040.Price1) AS Price1_Total" & _
                 " FROM WK_YPMF040 " & _
                 " GROUP BY WK_YPMF040.Key1, WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname" & _
                 " ORDER BY WK_YPMF040.Key1, WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname"
    End If
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    If wkRecordset.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = wkRecordset.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    
    Do While Not wkRecordset.EOF
        If chkRePrint.Value = 0 Then
            '出品者精算データオープン
            strSQL = "SELECT * FROM DT040" & _
                     " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND Pnum = " & wkRecordset.Fields("Pnum") & _
                     " ORDER BY Odate,Pnum,Num DESC"
            adoDT040.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If adoDT040.EOF = True Then
                adoDT040.AddNew
                intNum = 1
            Else
                intNum = CInt(adoDT040.Fields("Num")) + 1
                adoDT040.AddNew
            End If
            adoDT040.Fields("Odate") = wkRecordset.Fields("Odate")
            adoDT040.Fields("Pnum") = wkRecordset.Fields("Pnum")
            adoDT040.Fields("Num") = intNum
            adoDT040.Fields("Scode") = wkRecordset.Fields("Scode")
            adoDT040.Fields("Sname") = wkRecordset.Fields("Sname")
            
            'ワークから競売分の合計を計算
            strSQL = "SELECT Sum(WK_YPMF040.Price1) AS Kyobai_Price" & _
                     " FROM WK_YPMF040" & _
                     " WHERE WK_YPMF040.Div = 'A'" & _
                     " AND WK_YPMF040.Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND WK_YPMF040.Pnum = " & wkRecordset.Fields("Pnum") & _
                     " AND WK_YPMF040.Num = 1"
            wkRecordsetTemp.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
            If wkRecordsetTemp.EOF = False Then
                If IsNull(wkRecordsetTemp.Fields("Kyobai_Price")) = False Then
                    adoDT040.Fields("Total") = wkRecordsetTemp.Fields("Kyobai_Price")
                Else
                    adoDT040.Fields("Total") = 0
                End If
            Else
                adoDT040.Fields("Total") = 0
            End If
            wkRecordsetTemp.Close
            
            'ワークから注文分の合計を計算
            strSQL = "SELECT Sum(WK_YPMF040.Price1) AS Chumon_Price" & _
                     " FROM WK_YPMF040" & _
                     " WHERE WK_YPMF040.Div = 'B'" & _
                     " AND WK_YPMF040.Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND WK_YPMF040.Pnum = " & wkRecordset.Fields("Pnum") & _
                     " AND WK_YPMF040.Num = 1"
            wkRecordsetTemp.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
            If wkRecordsetTemp.EOF = False Then
                If IsNull(wkRecordsetTemp.Fields("Chumon_Price")) = False Then
                    adoDT040.Fields("Ototal") = wkRecordsetTemp.Fields("Chumon_Price")
                Else
                    adoDT040.Fields("Ototal") = 0
                End If
            Else
                adoDT040.Fields("Ototal") = 0
            End If
            wkRecordsetTemp.Close
            
            '金額＝競売金額＋注文金額
            curTotalPrice = CCur(adoDT040.Fields("Total")) + CCur(adoDT040.Fields("Ototal"))
            '手数料計算
            curBuff = curTotalPrice * curSrate / 100
            adoDT040.Fields("Charge") = Global_Rounding(curBuff, intSfraction, curSRounding)
            '消費税計算
            adoDT040.Fields("Tax") = Global_Get_Tax(curTotalPrice - CCur(adoDT040.Fields("Charge")), curTaxRate, intSfraction, curSRounding)
            adoDT040.Fields("Keep") = curSkeep
            '総合計＝合計－手数料＋消費税－維持管理費
            adoDT040.Fields("GTotal") = curTotalPrice - CCur(adoDT040.Fields("Charge")) + CCur(adoDT040.Fields("Tax")) - CCur(adoDT040.Fields("Keep"))
            If optFdiv(0).Value = True Then
                adoDT040.Fields("Pdiv") = PAYMENT_ON
                adoDT040.Fields("Pdate") = Format(Now(), "yyyy/mm/dd")
            Else
                adoDT040.Fields("Pdiv") = PAYMENT_OFF
                adoDT040.Fields("Pdate") = Null
            End If
            adoDT040.Fields("Pay") = intR
            adoDT040.Fields("Itime") = Format(Now(), "hh:mm:ss")
            adoDT040.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
            adoDT040.Update
            
            'ワークの更新
            strSQL = "UPDATE WK_YPMF040" & _
                     " SET WK_YPMF040.Total = " & curTotalPrice & "," & _
                     " WK_YPMF040.Charge = " & adoDT040.Fields("Charge") & "," & _
                     " WK_YPMF040.Tax = " & adoDT040.Fields("Tax") & "," & _
                     " WK_YPMF040.Keep = " & adoDT040.Fields("Keep") & "," & _
                     " WK_YPMF040.GTotal = " & adoDT040.Fields("GTotal") & _
                     " WHERE WK_YPMF040.Odate = '" & adoDT040.Fields("Odate") & "'" & _
                     " AND WK_YPMF040.Pnum = " & adoDT040.Fields("Pnum")
            g_clsAdoAccess.Connection.Execute strSQL
        
            adoDT040.Close
        Else
            '出品者精算データオープン
            strSQL = "SELECT * FROM DT040" & _
                     " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND Pnum = " & wkRecordset.Fields("Pnum") & _
                     " AND Num = " & Mid(wkRecordset.Fields("Key1"), 5, 2) & _
                     " ORDER BY Odate,Pnum,Num"
            adoDT040.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If adoDT040.EOF = False Then
                '出品者精算データ更新
                
                'ワークから競売分の合計を計算
                strSQL = "SELECT Sum(WK_YPMF040.Price1) AS Kyobai_Price" & _
                         " FROM WK_YPMF040" & _
                         " WHERE WK_YPMF040.Div = 'A'" & _
                         " AND WK_YPMF040.Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                         " AND WK_YPMF040.Pnum = " & wkRecordset.Fields("Pnum") & _
                         " AND WK_YPMF040.Num = " & Mid(wkRecordset.Fields("Key1"), 5, 2)
                wkRecordsetTemp.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
                If wkRecordsetTemp.EOF = False Then
                    If IsNull(wkRecordsetTemp.Fields("Kyobai_Price")) = False Then
                        adoDT040.Fields("Total") = wkRecordsetTemp.Fields("Kyobai_Price")
                    Else
                        adoDT040.Fields("Total") = 0
                    End If
                Else
                    adoDT040.Fields("Total") = 0
                End If
                wkRecordsetTemp.Close
                
                'ワークから注文分の合計を計算
                strSQL = "SELECT Sum(WK_YPMF040.Price1) AS Chumon_Price" & _
                         " FROM WK_YPMF040" & _
                         " WHERE WK_YPMF040.Div = 'B'" & _
                         " AND WK_YPMF040.Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                         " AND WK_YPMF040.Pnum = " & wkRecordset.Fields("Pnum") & _
                         " AND WK_YPMF040.Num = " & Mid(wkRecordset.Fields("Key1"), 5, 2)
                wkRecordsetTemp.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
                If wkRecordsetTemp.EOF = False Then
                    If IsNull(wkRecordsetTemp.Fields("Chumon_Price")) = False Then
                        adoDT040.Fields("Ototal") = wkRecordsetTemp.Fields("Chumon_Price")
                    Else
                        adoDT040.Fields("Ototal") = 0
                    End If
                Else
                    adoDT040.Fields("Ototal") = 0
                End If
                wkRecordsetTemp.Close
                
                '金額＝競売金額＋注文金額
                curTotalPrice = CCur(adoDT040.Fields("Total")) + CCur(adoDT040.Fields("Ototal"))
                '手数料計算
                curBuff = curTotalPrice * curSrate / 100
                adoDT040.Fields("Charge") = Global_Rounding(curBuff, intSfraction, curSRounding)
                '消費税計算
                adoDT040.Fields("Tax") = Global_Get_Tax(curTotalPrice - CCur(adoDT040.Fields("Charge")), curTaxRate, intSfraction, curSRounding)
                adoDT040.Fields("Keep") = curSkeep
                '総合計＝合計－手数料＋消費税－維持管理費
                adoDT040.Fields("GTotal") = curTotalPrice - CCur(adoDT040.Fields("Charge")) + CCur(adoDT040.Fields("Tax")) - CCur(adoDT040.Fields("Keep"))
                If optFdiv(0).Value = True Then
                    adoDT040.Fields("Pdiv") = PAYMENT_ON
                    adoDT040.Fields("Pdate") = Format(Now(), "yyyy/mm/dd")
                Else
                    adoDT040.Fields("Pdiv") = PAYMENT_OFF
                    adoDT040.Fields("Pdate") = Null
                End If
                adoDT040.Fields("Pay") = intR

                adoDT040.Update
            
                '合計金額
                curTotalPrice = CCur(adoDT040.Fields("Total")) + CCur(adoDT040.Fields("Ototal"))
                
                'ワークの更新
                strSQL = "UPDATE WK_YPMF040" & _
                         " SET WK_YPMF040.Total = " & curTotalPrice & "," & _
                         " WK_YPMF040.Charge = " & adoDT040.Fields("Charge") & "," & _
                         " WK_YPMF040.Tax = " & adoDT040.Fields("Tax") & "," & _
                         " WK_YPMF040.Keep = " & adoDT040.Fields("Keep") & "," & _
                         " WK_YPMF040.GTotal = " & adoDT040.Fields("GTotal") & _
                         " WHERE WK_YPMF040.Odate = '" & adoDT040.Fields("Odate") & "'" & _
                         " AND WK_YPMF040.Pnum = " & adoDT040.Fields("Pnum") & _
                         " AND WK_YPMF040.Key1 = '" & wkRecordset.Fields("Key1") & "'"
                g_clsAdoAccess.Connection.Execute strSQL
            End If
            adoDT040.Close
        End If
        
        wkRecordset.MoveNext
        lngCount = lngCount + 1
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Old_Cancel:
    Loop
    
    wkRecordset.Requery
    wkRecordset.Close
    
    g_clsAdoSQL.Connection.CommitTrans
    
    If lngCount = 0 Then
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Old_Exit:
    End If
    
    MakePrintWork_Old = True
    
MakePrintWork_Old_Exit:
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    
    Exit Function

MakePrintWork_Old_Cancel:

    g_clsAdoSQL.Connection.RollbackTrans
    GoTo MakePrintWork_Old_Exit:

MakePrintWork_Old_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    MakePrintWork_Old = False
    Call MsgBox("印刷ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Old_Err")
    GoTo MakePrintWork_Old_Exit:

End Function

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２２
'更新履歴　：
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
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboPnum_Err")

End Sub

Private Sub imnRePrintNum_GotFocus(Index As Integer)

    imnRePrintNum(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRePrintNum_LostFocus(Index As Integer)

    imnRePrintNum(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cboPnum(0).SetFocus

End Sub

Private Sub optFdiv_Click(Index As Integer)

    Dim intIndex1 As Integer

    On Error GoTo optFdiv_Click_Err

    '背景色の変更
    For intIndex1 = 0 To optFdiv.Count - 1
        If intIndex1 = Index Then
            optFdiv(intIndex1).BackColor = BUTTON_ON
        Else
            optFdiv(intIndex1).BackColor = BUTTON_OFF
        End If
    Next intIndex1

    Exit Sub

optFdiv_Click_Err:

    Call MsgBox("入金種別クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "optFdiv_Click_Err")

End Sub

'目　的　　：ActiveReportの印刷
'条　件　　：
'結　果　　：
'引　数　　：0:プレビュー 1:印刷
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２２
'更新履歴　：
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf040
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "出品者精算明細票"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .Caption = "出品者精算明細票"
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

'目　的　　：印刷用ワーク作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／２２
'更新履歴　：
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim adoDT010 As New ADODB.Recordset
    Dim adoDT011 As New ADODB.Recordset
    Dim adoDT021 As New ADODB.Recordset
    Dim adoDT030 As New ADODB.Recordset
    Dim adoDT040 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim wkRecordsetTemp As New ADODB.Recordset
    Dim strBuff1 As String

    Dim intLine As Integer                  '行番号
    Dim lngCount As Long                    'レコード件数
    Dim curSkeep As Currency                '出品者維持管理費(標準)
    Dim curSkeepCurrent As Currency         '出品者維持管理費(今回)
    Dim curSrate As Currency                '出品者手数料率
    Dim intSfraction As Integer             '出品者端数処理
    Dim intNum As Integer                   '回数
    Dim curTaxRate As Currency              '消費税率
    Dim intR As Integer                     '支払種別
    Dim curSRounding As Currency            '出品者丸め単位
    Dim curTotalPrice As Currency           '金額(競売金額＋注文金額)
    Dim varKey1 As Variant
    Dim curBuff As Currency
    Dim blnFlg As Boolean
    Dim strOdateNum As String
    Dim strSoukinMsg As String

    Dim curRrate As Currency                '201107 出品者受付伝票代
    Dim curEf As Currency                   '201107 出品者絵札代
    
    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
'********** 初期処理 **********
    
    strOdateNum = Global_Get_NumericDay(lblOdate.Caption)
    
    '初期化
    curSkeep = 0        '出品者維持管理費
    intSfraction = 0    '出品者端数処理
    curSrate = 0        '出品者手数料率
    curTaxRate = 0      '消費税率
    intR = 0            '支払種別
    curSRounding = 0    '出品者丸め単位
    
    curRrate = 0        '201107 出品者受付伝票代
    curEf = 0           '201107 出品者絵札代
    
    If optFdiv(0).Value = True Then
        intR = PAYMENT_DIV_CASH
    ElseIf optFdiv(1).Value = True Then
        intR = PAYMENT_DIV_CHECK
    ElseIf optFdiv(2).Value = True Then
        intR = PAYMENT_DIV_TRANSFER
    End If
    
    '設定マスタオープン
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        If Not IsNull(adoMT010.Fields("Skeep")) Then curSkeep = adoMT010.Fields("Skeep")
        If Not IsNull(adoMT010.Fields("Sfraction")) Then intSfraction = adoMT010.Fields("Sfraction")
        If Not IsNull(adoMT010.Fields("Srate")) Then curSrate = adoMT010.Fields("Srate")
        If Not IsNull(adoMT010.Fields("SRounding")) Then curSRounding = adoMT010.Fields("SRounding")
        '201107
        If Not IsNull(adoMT010.Fields("Rrate")) Then curRrate = adoMT010.Fields("Rrate")
        If Not IsNull(adoMT010.Fields("Ef")) Then curEf = adoMT010.Fields("Ef")
    End If
    adoMT010.Close
    
    '消費税率取得
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, Trim(lblOdate.Caption))
    
'********** ワーク **********
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF040"
    g_clsAdoAccess.Connection.Execute strSQL

    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF040"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '受付データオープン
    strSQL = "{call sp_YPMF0401;1('" & Trim(lblOdate.Caption) & "'," & _
              cboPnum(0).Text & "," & cboPnum(1).Text & ")}"
    adoDT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT010.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = adoDT010.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    
    Do While Not adoDT010.EOF
'        '受付明細データオープン
'        strSQL = "{call sp_YPMF0402;1('" & lblOdate.Caption & "'," & adoDT010.Fields("Pnum") & ")}"
'        adoDT011.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
'        Do While Not adoDT011.EOF
'            If IsNull(adoDT011.Fields("Price")) = False And adoDT011.Fields("Price") <> 0 Then
'                '********** 注文分 **********
'                wkRecordset.AddNew
'                wkRecordset.Fields("Key1") = Format(adoDT010.Fields("Pnum"), "0000") & "01"
'                wkRecordset.Fields("Key2") = Format(adoDT010.Fields("Pnum"), "0000") & "01"
'                wkRecordset.Fields("Num") = 1
'                wkRecordset.Fields("Div") = "B"
'                wkRecordset.Fields("Odate") = Trim(lblOdate.Caption)
'                wkRecordset.Fields("Pnum") = adoDT010.Fields("Pnum")
'                wkRecordset.Fields("Scode") = adoDT010.Fields("Scode")
'                wkRecordset.Fields("Sname") = adoDT010.Fields("Sname")
'                wkRecordset.Fields("Line") = adoDT011.Fields("Line")
'                wkRecordset.Fields("Icode") = adoDT011.Fields("Icode")
'                wkRecordset.Fields("Iname") = adoDT011.Fields("Iname")
'                wkRecordset.Fields("Qty") = adoDT011.Fields("Qty")
'                wkRecordset.Fields("Price1") = adoDT011.Fields("Price")
'                wkRecordset.Fields("Bcode") = adoDT011.Fields("Bcode")
'                wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, adoDT011.Fields("Bcode"), Trim(lblOdate.Caption), strBuff1)
''                wkRecordset.Fields("Total") = 0
''                wkRecordset.Fields("Charge") = 0
''                wkRecordset.Fields("Tax") = 0
''                wkRecordset.Fields("Keep") = curSkeep
''                wkRecordset.Fields("GTotal") = 0
'                wkRecordset.Fields("Itime") = Format(Now(), "hh時mm分")
'                wkRecordset.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
'                wkRecordset.Fields("Pname") = Trim(lblPname.Caption)
'                wkRecordset.Fields("Idiv") = 0
'                wkRecordset.Fields("Ocode") = Format(adoDT010.Fields("Pnum"), "0000") & "*"
'                wkRecordset.Fields("RePrint") = adoDT011.Fields("Sdiv")
'                wkRecordset.Update
'
'                '受付データ更新
'                adoDT011.Fields("Sdiv") = EXHIBITION_REPORT_ON
'                adoDT011.Fields("Snum") = 1
'                adoDT011.Update
'            End If
'            adoDT011.MoveNext
'        Loop
'        adoDT011.Close
        
        '競売明細データオープン
        strSQL = "{call sp_YPMF0403;1('" & strOdateNum & "'," & adoDT010.Fields("Pnum") & ")}"
        adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        Do While Not adoDT021.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Key1") = Format(adoDT010.Fields("Pnum"), "0000") & "01"
            wkRecordset.Fields("Key2") = Format(adoDT010.Fields("Pnum"), "0000") & "01"
            wkRecordset.Fields("Num") = 1
            wkRecordset.Fields("Div") = "A"
            wkRecordset.Fields("Odate") = Trim(lblOdate.Caption)
            wkRecordset.Fields("Pnum") = adoDT010.Fields("Pnum")
            wkRecordset.Fields("Scode") = adoDT010.Fields("Scode")
            wkRecordset.Fields("Sname") = Trim(adoDT010.Fields("Sname")) & "　様"
            wkRecordset.Fields("Line") = adoDT021.Fields("PnumLine")
            wkRecordset.Fields("Icode") = adoDT021.Fields("Icode")
            wkRecordset.Fields("Iname") = adoDT021.Fields("Iname")
            wkRecordset.Fields("Qty") = adoDT021.Fields("Qty")
            wkRecordset.Fields("Price1") = adoDT021.Fields("Price")
            wkRecordset.Fields("Bcode") = adoDT021.Fields("Bcode")
            If Not IsNull(wkRecordset.Fields("Bcode")) Then
                wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, adoDT021.Fields("Bcode"), Trim(lblOdate.Caption), strBuff1)
            Else
                wkRecordset.Fields("Bname") = ""
            End If
'            wkRecordset.Fields("Total") = 0
'            wkRecordset.Fields("Charge") = 0
'            wkRecordset.Fields("Tax") = 0
'            wkRecordset.Fields("Keep") = curSkeep
'            wkRecordset.Fields("GTotal") = 0
            wkRecordset.Fields("Itime") = Format(Now(), "hh時mm分")
            wkRecordset.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
            wkRecordset.Fields("Pname") = Trim(lblPname.Caption)
            If IsNull(adoDT021.Fields("Sline")) = False Then
                wkRecordset.Fields("Idiv") = adoDT021.Fields("Sline")
            Else
                wkRecordset.Fields("Idiv") = 0
            End If
            If IsNull(adoDT021.Fields("Idiv")) = False Then
                wkRecordset.Fields("Result") = adoDT021.Fields("Idiv")
            Else
                wkRecordset.Fields("Result") = 0
            End If
            wkRecordset.Fields("Ocode") = right$(adoDT021.Fields("Ocode"), 4)
            wkRecordset.Fields("RePrint") = adoDT021.Fields("Sdiv")

'202308 出品者登録番号
            wkRecordset.Fields("Pname") = Trim(adoDT010.Fields("Addres"))
'202308 出品者登録番号
            
            wkRecordset.Update
        
            '競売明細データ更新
            adoDT021.Fields("Sdiv") = EXHIBITION_REPORT_ON
            adoDT021.Fields("Snum") = 1
            adoDT021.Update
            adoDT021.MoveNext
        Loop
        adoDT021.Close
        
        adoDT010.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    adoDT010.Close
    wkRecordset.Close
    
'********** 出品者精算データ作成(再発行分も含む) **********

    'ワークオープン
    strSQL = "SELECT WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname, Sum(WK_YPMF040.Price1) AS Price1_Total" & _
             " FROM WK_YPMF040 " & _
             " GROUP BY WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname" & _
             " ORDER BY WK_YPMF040.Odate, WK_YPMF040.Pnum, WK_YPMF040.Scode, WK_YPMF040.Sname"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    If wkRecordset.EOF = False Then
        frmCount.fpProgressBar1.Value = 0
        frmCount.fpProgressBar1.Max = wkRecordset.RecordCount
        frmCount.Show
        Me.Enabled = False
    End If
    
    lngCount = 0
    Do While Not wkRecordset.EOF
        '出品者精算データオープン
        strSQL = "SELECT * FROM DT040" & _
                 " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                 " AND Pnum = " & wkRecordset.Fields("Pnum") & _
                 " AND Num = 1 " & _
                 " ORDER BY Odate,Pnum,Num DESC"
        adoDT040.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If adoDT040.EOF = True Then
            adoDT040.AddNew
        End If
        adoDT040.Fields("Odate") = wkRecordset.Fields("Odate")
        adoDT040.Fields("Pnum") = wkRecordset.Fields("Pnum")
        adoDT040.Fields("Num") = 1
        adoDT040.Fields("Scode") = wkRecordset.Fields("Scode")
        adoDT040.Fields("Sname") = wkRecordset.Fields("Sname")
        
        'ワークから競売分の合計を計算
        strSQL = "SELECT Sum(WK_YPMF040.Price1) AS Kyobai_Price" & _
                 " FROM WK_YPMF040" & _
                 " WHERE WK_YPMF040.Div = 'A'" & _
                 " AND WK_YPMF040.Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                 " AND WK_YPMF040.Pnum = " & wkRecordset.Fields("Pnum") & _
                 " AND WK_YPMF040.Num = 1"
        wkRecordsetTemp.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
        If wkRecordsetTemp.EOF = False Then
            If IsNull(wkRecordsetTemp.Fields("Kyobai_Price")) = False Then
                adoDT040.Fields("Total") = wkRecordsetTemp.Fields("Kyobai_Price")
            Else
                adoDT040.Fields("Total") = 0
            End If
        Else
            adoDT040.Fields("Total") = 0
        End If
        wkRecordsetTemp.Close
        
        'ワークから注文分の合計を計算
        strSQL = "SELECT Sum(WK_YPMF040.Price1) AS Chumon_Price" & _
                 " FROM WK_YPMF040" & _
                 " WHERE WK_YPMF040.Div = 'B'" & _
                 " AND WK_YPMF040.Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                 " AND WK_YPMF040.Pnum = " & wkRecordset.Fields("Pnum") & _
                 " AND WK_YPMF040.Num = 1"
        wkRecordsetTemp.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
        If wkRecordsetTemp.EOF = False Then
            If IsNull(wkRecordsetTemp.Fields("Chumon_Price")) = False Then
                adoDT040.Fields("Ototal") = wkRecordsetTemp.Fields("Chumon_Price")
            Else
                adoDT040.Fields("Ototal") = 0
            End If
        Else
            adoDT040.Fields("Ototal") = 0
        End If
        wkRecordsetTemp.Close
        
        
        '201107 受付明細から受付件数を計算
        strSQL = "SELECT COUNT(*) AS Ef_Count" & _
                 " FROM DT011" & _
                 " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                 " AND Pnum = " & wkRecordset.Fields("Pnum")
        wkRecordsetTemp.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If wkRecordsetTemp.EOF = False Then
            If IsNull(wkRecordsetTemp.Fields("Ef_Count")) = False Then
                If (wkRecordsetTemp.Fields("Ef_Count") = 0) Then
                    adoDT040.Fields("Rrate") = 0
                    adoDT040.Fields("Ef") = 0
                Else
                    Dim cnt As Double
                    cnt = wkRecordsetTemp.Fields("Ef_Count")
                    
                    If cnt <= 20 Then
                    adoDT040.Fields("Rrate") = CCur(curRrate)
                    Else
                    adoDT040.Fields("Rrate") = CCur(ToRoundDown(cnt / 20, 0) * curRrate)
                    End If
                    adoDT040.Fields("Ef") = CCur(wkRecordsetTemp.Fields("Ef_Count") * curEf)
                End If
            Else
                adoDT040.Fields("Rrate") = 0
                adoDT040.Fields("Ef") = 0
            End If
        Else
            adoDT040.Fields("Rrate") = 0
            adoDT040.Fields("Ef") = 0
        End If
        wkRecordsetTemp.Close
        '201107

        '金額＝競売金額＋注文金額
        curTotalPrice = CCur(adoDT040.Fields("Total")) + CCur(adoDT040.Fields("Ototal"))
        '手数料計算
        curBuff = curTotalPrice * curSrate / 100
        adoDT040.Fields("Charge") = Global_Rounding(curBuff, intSfraction, curSRounding)
        '202308 税抜き維持管理費
        adoDT040.Fields("Keep") = Global_Rounding(curSkeep / (1 + (curTaxRate / 100)), intSfraction, 1)
        '202308 税抜き受付伝票
        adoDT040.Fields("Rrate") = Global_Rounding(adoDT040.Fields("Rrate") / (1 + (curTaxRate / 100)), intSfraction, 1)
        '202308 税抜き荷札
        adoDT040.Fields("Ef") = Global_Rounding(adoDT040.Fields("Ef") / (1 + (curTaxRate / 100)), intSfraction, 1)
        
        '202308 消費税計算　競売金額＋注文金額-手数料-税抜き維持管理費-税抜き受付伝票代-税抜き絵札代　から計算
        adoDT040.Fields("Tax") = Global_Get_Tax(curTotalPrice - CCur(adoDT040.Fields("Charge")) - CCur(adoDT040.Fields("Keep")) - CCur(adoDT040.Fields("Rrate")) - CCur(adoDT040.Fields("Ef")), curTaxRate, intSfraction, 1)
        
        Dim curStamp As Currency
        '201107 印紙代金を計算 競売金額＋注文金額 - 手数料を元に　strSQL = "SELECT Price AS Stamp" & _
                 " FROM MT080" & _
                 " WHERE " & (curTotalPrice - CCur(adoDT040.Fields("Charge"))) & ">=SLimit " & _
                 " AND FLimit>= " & (curTotalPrice - CCur(adoDT040.Fields("Charge")))
        '202308 印紙代金を計算 競売金額 ＋ 注文金額 - 手数料 - 税抜き維持管理費 - 税抜き受付伝票代 - 税抜き絵札代 を元に
        strSQL = "SELECT Price AS Stamp" & _
                 " FROM MT080" & _
                 " WHERE " & (Global_Rounding(curTotalPrice - CCur(adoDT040.Fields("Charge") - CCur(adoDT040.Fields("Keep")) - CCur(adoDT040.Fields("Rrate")) - CCur(adoDT040.Fields("Ef")) / (1 + (curTaxRate / 100))), intSfraction, 1)) & ">=SLimit " & _
                 " AND FLimit>= " & (Global_Rounding(curTotalPrice - CCur(adoDT040.Fields("Charge") - CCur(adoDT040.Fields("Keep")) - CCur(adoDT040.Fields("Rrate")) - CCur(adoDT040.Fields("Ef")) / (1 + (curTaxRate / 100))), intSfraction, 1))
       
        wkRecordsetTemp.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        
        If wkRecordsetTemp.EOF = False Then
            If IsNull(wkRecordsetTemp.Fields("Stamp")) = False Then
                curStamp = CCur(wkRecordsetTemp.Fields("Stamp"))
            Else
                curStamp = 0
            End If
        Else
            curStamp = 0
        End If
        wkRecordsetTemp.Close
        adoDT040.Fields("Stamp") = curStamp
        '201107
        
        
        '総合計＝合計－手数料＋消費税－維持管理費
        'adoDT040.Fields("GTotal") = curTotalPrice - CCur(adoDT040.Fields("Charge")) + CCur(adoDT040.Fields("Tax")) - CCur(adoDT040.Fields("Keep"))
        '201107 総合計＝合計－手数料＋消費税－維持管理費-受付伝票代-絵札代-印紙代金
        'adoDT040.Fields("GTotal") = curTotalPrice - CCur(adoDT040.Fields("Charge")) + CCur(adoDT040.Fields("Tax")) - CCur(adoDT040.Fields("Keep")) - CCur(adoDT040.Fields("Rrate")) - CCur(adoDT040.Fields("Ef")) - CCur(adoDT040.Fields("Stamp"))
        
        '202308 総合計＝合計＋消費税-印紙代金
        adoDT040.Fields("GTotal") = curTotalPrice - CCur(adoDT040.Fields("Charge")) - CCur(adoDT040.Fields("Keep")) - CCur(adoDT040.Fields("Rrate")) - CCur(adoDT040.Fields("Ef")) - CCur(adoDT040.Fields("Stamp")) + adoDT040.Fields("Tax")

        
        If optFdiv(0).Value = True Then
            adoDT040.Fields("Pdiv") = PAYMENT_ON
            adoDT040.Fields("Pdate") = Format(Now(), "yyyy/mm/dd")
        Else
            adoDT040.Fields("Pdiv") = PAYMENT_OFF
            adoDT040.Fields("Pdate") = Null
        End If
        adoDT040.Fields("Pay") = intR
        adoDT040.Fields("Itime") = Format(Now(), "hh:mm:ss")
        adoDT040.Fields("Pcode") = IIf(IsNumeric(cboPcode.Text), cboPcode.Text, Null)
        adoDT040.Update
        
        '新規分のみの場合
        If chkRePrint.Value = 0 Then
            'ワークから未発行分があるかチェック
            strSQL = "SELECT * FROM WK_YPMF040" & _
                     " WHERE WK_YPMF040.Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                     " AND WK_YPMF040.Pnum = " & wkRecordset.Fields("Pnum") & _
                     " AND WK_YPMF040.Num = 1" & _
                     " AND (WK_YPMF040.RePrint = 0 OR WK_YPMF040.RePrint IS NULL)"
            wkRecordsetTemp.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
            If wkRecordsetTemp.EOF = True Then
                'ワーク削除
                strSQL = "DELETE FROM WK_YPMF040" & _
                         " WHERE WK_YPMF040.Odate = '" & adoDT040.Fields("Odate") & "'" & _
                         " AND WK_YPMF040.Pnum = " & adoDT040.Fields("Pnum")
                g_clsAdoAccess.Connection.Execute strSQL
            End If
            wkRecordsetTemp.Close
        End If
        
        '********** 送金 **********
        
        strSoukinMsg = ""
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & wkRecordset.Fields("Odate") & "'" & _
                 " AND Pnum = " & wkRecordset.Fields("Pnum")
        adoDT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoDT010.EOF = False Then
            If Not IsNull(adoDT010.Fields("Soukin")) And adoDT010.Fields("Soukin") = 1 Then
                strSoukinMsg = "※送金"
            End If
        End If
        adoDT010.Close
        
        
        'ワークの更新
        strSQL = "UPDATE WK_YPMF040" & _
                 " SET WK_YPMF040.Total = " & curTotalPrice & "," & _
                 " WK_YPMF040.Charge = " & adoDT040.Fields("Charge") & "," & _
                 " WK_YPMF040.Tax = " & adoDT040.Fields("Tax") & "," & _
                 " WK_YPMF040.Keep = " & adoDT040.Fields("Keep") & "," & _
                 " WK_YPMF040.GTotal = " & adoDT040.Fields("GTotal") & "," & _
                 " WK_YPMF040.Rrate = " & adoDT040.Fields("Rrate") & "," & _
                 " WK_YPMF040.Ef = " & adoDT040.Fields("Ef") & "," & _
                 " WK_YPMF040.Stamp = " & adoDT040.Fields("Stamp") & "," & _
                 " WK_YPMF040.Soukin = '" & strSoukinMsg & "'" & _
                 " WHERE WK_YPMF040.Odate = '" & adoDT040.Fields("Odate") & "'" & _
                 " AND WK_YPMF040.Pnum = " & adoDT040.Fields("Pnum")
        g_clsAdoAccess.Connection.Execute strSQL
    
        adoDT040.Close
        
        wkRecordset.MoveNext
        lngCount = lngCount + 1
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    
    wkRecordset.Requery
    wkRecordset.Close
    
    g_clsAdoSQL.Connection.CommitTrans
    
    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF040"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    If wkRecordset.EOF = True Then
        wkRecordset.Close
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If
    wkRecordset.Close
    
    MakePrintWork = True
    
MakePrintWork_Exit:
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    
    Exit Function

MakePrintWork_Cancel:

    g_clsAdoSQL.Connection.RollbackTrans
    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    MakePrintWork = False
    Call MsgBox("印刷ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function
Public Function ToRoundDown(ByVal dValue As Double, ByVal iDigits As Integer) As Double
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundDown = Int(dValue * dCoef) / dCoef
        Case Is < 0
            ToRoundDown = Fix(dValue * dCoef) / dCoef
        Case Else
            ToRoundDown = dValue
    End Select
End Function
