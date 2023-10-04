VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmYpmf010 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   10155
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13065
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf010.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   13065
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraMeisai 
      Height          =   6195
      Left            =   60
      TabIndex        =   37
      Top             =   3180
      Width           =   12915
      Begin VB.CheckBox chkUpdflg 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   98
         Top             =   180
         Width           =   615
      End
      Begin VB.Frame fraDetail 
         Height          =   2355
         Left            =   9840
         TabIndex        =   73
         Top             =   120
         Visible         =   0   'False
         Width           =   2955
         Begin VB.CheckBox chkSdiv 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   86
            Top             =   180
            Width           =   615
         End
         Begin VB.CheckBox chkBdiv 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   85
            Top             =   600
            Width           =   615
         End
         Begin VB.CheckBox chkIdiv 
            Height          =   435
            Left            =   2280
            TabIndex        =   74
            Top             =   1860
            Width           =   435
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   87
            Top             =   180
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "出品伝票出力済み"
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
            LabelWidth      =   120
            LabelHeight     =   25
            LabelLeft       =   8
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
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "買主伝票出力済み"
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
            LabelWidth      =   120
            LabelHeight     =   25
            LabelLeft       =   8
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
            Index           =   18
            Left            =   120
            TabIndex        =   89
            Top             =   1020
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "買主精算回数"
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
            LabelLeft       =   23
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
            Left            =   120
            TabIndex        =   90
            Top             =   1440
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "出品者精算回数"
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
            LabelLeft       =   16
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
         Begin imNumber6Ctl.imNumber imnBnum 
            Height          =   375
            Left            =   2220
            TabIndex        =   91
            Top             =   1020
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   661
            Calculator      =   "frmYpmf010.frx":0CFA
            Caption         =   "frmYpmf010.frx":0D1A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf010.frx":0D88
            Keys            =   "frmYpmf010.frx":0DA6
            Spin            =   "frmYpmf010.frx":0DF0
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
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnSnum 
            Height          =   375
            Left            =   2220
            TabIndex        =   92
            Top             =   1440
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   661
            Calculator      =   "frmYpmf010.frx":0E18
            Caption         =   "frmYpmf010.frx":0E38
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf010.frx":0EA6
            Keys            =   "frmYpmf010.frx":0EC4
            Spin            =   "frmYpmf010.frx":0F0E
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
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   93
            Top             =   1860
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "入力済みフラグ"
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
            LabelWidth      =   97
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
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   2340
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmYpmf010.frx":0F36
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "明細挿入(&I)"
         Height          =   375
         Left            =   9360
         TabIndex        =   29
         Top             =   2100
         Width           =   1575
      End
      Begin VB.CommandButton cmdPast 
         Caption         =   "明細貼付(&P)"
         Height          =   375
         Left            =   7800
         TabIndex        =   28
         Top             =   2100
         Width           =   1575
      End
      Begin VB.CheckBox chkChumon 
         Caption         =   "注 文 分(&O)"
         Height          =   435
         Left            =   3060
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "..."
         Height          =   375
         Left            =   2040
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "明細コピー(&C)"
         Height          =   375
         Left            =   6240
         TabIndex        =   27
         Top             =   2100
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear_Dst 
         Caption         =   "明細クリア(&N)"
         Height          =   375
         Left            =   4680
         TabIndex        =   26
         Top             =   2100
         Width           =   1575
      End
      Begin VB.CommandButton cmdMt050 
         Caption         =   "植木のマスタ登録(F11)"
         Height          =   375
         Left            =   6660
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   600
         Width           =   2415
      End
      Begin CSComboLib.CSComboBox cboIcode 
         Height          =   360
         Left            =   9000
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         Enabled         =   0   'False
         ColDelim        =   ";"
         ColWidths       =   "2;20"
         Contents        =   "frmYpmf010.frx":1090
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   5
         Text            =   "99999"
      End
      Begin MSComctlLib.ListView lsvMeisai 
         Height          =   3135
         Left            =   180
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2520
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "行"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "コード"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "植木名称"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "数　量"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "済"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "単　価"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "下代単価"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "売立金額"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "買主"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "買主名称"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "出品者伝票区分"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "買主伝票区分"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "買主精算回数"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "出品者精算回数"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "変更フラグ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "明細削除(&D)"
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   2100
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "明細登録(&A)"
         Height          =   375
         Left            =   1560
         TabIndex        =   24
         Top             =   2100
         Width           =   1575
      End
      Begin imNumber6Ctl.imNumber imnQty 
         Height          =   435
         Left            =   1560
         TabIndex        =   18
         Top             =   1560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   767
         Calculator      =   "frmYpmf010.frx":10A9
         Caption         =   "frmYpmf010.frx":10C9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":1137
         Keys            =   "frmYpmf010.frx":1155
         Spin            =   "frmYpmf010.frx":119F
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
         ValueVT         =   5
         Value           =   999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imText6Ctl.imText txtIname 
         Height          =   405
         Left            =   1560
         TabIndex        =   17
         Top             =   1080
         Width           =   8595
         _Version        =   65536
         _ExtentX        =   15161
         _ExtentY        =   714
         Caption         =   "frmYpmf010.frx":11C7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":1235
         Key             =   "frmYpmf010.frx":1253
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
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
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
         TabIndex        =   44
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "行番号"
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
         Index           =   0
         Left            =   60
         TabIndex        =   54
         Top             =   1560
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "数　量"
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
      Begin imNumber6Ctl.imNumber imnNo 
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   180
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmYpmf010.frx":1297
         Caption         =   "frmYpmf010.frx":12B7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":1325
         Keys            =   "frmYpmf010.frx":1343
         Spin            =   "frmYpmf010.frx":138D
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
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0"
         HighlightText   =   0
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
         ShowContextMenu =   -1
         ValueVT         =   2011496453
         Value           =   99
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSComboLib.CSComboBox cboIcode_Kana 
         Height          =   405
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   4935
         _Version        =   262145
         _ExtentX        =   8705
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "10;20;4"
         Contents        =   "frmYpmf010.frx":13B5
         Extended        =   -1  'True
         ListBoxWidth    =   700
         MaxLength       =   20
         Text            =   "WWWWWWWWWWWWWWWWWWWQ"
         ValueCol        =   2
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   9
         Left            =   60
         TabIndex        =   62
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "植木ｶﾅ検索"
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
         LabelWidth      =   77
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
      Begin imNumber6Ctl.imNumber imnQty_Total 
         Height          =   375
         Left            =   6420
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5700
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   661
         Calculator      =   "frmYpmf010.frx":13CE
         Caption         =   "frmYpmf010.frx":13EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":145C
         Keys            =   "frmYpmf010.frx":147A
         Spin            =   "frmYpmf010.frx":14BC
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2011496453
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   11
         Left            =   60
         TabIndex        =   69
         Top             =   1080
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "植木名"
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
      Begin VB.Frame fraChumon 
         BorderStyle     =   0  'なし
         Height          =   1095
         Left            =   4440
         TabIndex        =   79
         Top             =   1380
         Width           =   8415
         Begin imNumber6Ctl.imNumber imnPrice1 
            Height          =   435
            Left            =   1080
            TabIndex        =   20
            Top             =   180
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   767
            Calculator      =   "frmYpmf010.frx":14E4
            Caption         =   "frmYpmf010.frx":1504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf010.frx":1572
            Keys            =   "frmYpmf010.frx":1590
            Spin            =   "frmYpmf010.frx":15DA
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   9999999
            MinValue        =   -9999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   9999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   13
            Left            =   60
            TabIndex        =   80
            Top             =   180
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "単　価"
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
         Begin imNumber6Ctl.imNumber imnPrice 
            Height          =   435
            Left            =   6720
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   767
            Calculator      =   "frmYpmf010.frx":1602
            Caption         =   "frmYpmf010.frx":1622
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf010.frx":1690
            Keys            =   "frmYpmf010.frx":16AE
            Spin            =   "frmYpmf010.frx":16F8
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
            Index           =   14
            Left            =   5460
            TabIndex        =   81
            Top             =   180
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "売立金額"
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
         Begin CSComboLib.CSComboBox cboBcode 
            Height          =   435
            Left            =   1080
            TabIndex        =   23
            Top             =   660
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
            _ExtentY        =   767
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColDelim        =   ";"
            ColWidths       =   "4;20"
            Contents        =   "frmYpmf010.frx":1720
            Extended        =   -1  'True
            ListBoxWidth    =   500
            MaxLength       =   4
            Text            =   "9999"
            ValueCol        =   0
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   15
            Left            =   60
            TabIndex        =   82
            Top             =   660
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "買　主"
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
         Begin imNumber6Ctl.imNumber imnPrice2 
            Height          =   435
            Left            =   3900
            TabIndex        =   21
            Top             =   180
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   767
            Calculator      =   "frmYpmf010.frx":1739
            Caption         =   "frmYpmf010.frx":1759
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf010.frx":17C7
            Keys            =   "frmYpmf010.frx":17E5
            Spin            =   "frmYpmf010.frx":182F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   9999999
            MinValue        =   -9999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   9999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   21
            Left            =   2700
            TabIndex        =   94
            Top             =   180
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "仕入単価"
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
            LabelLeft       =   8
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
         Begin VB.Label lblBname 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000005&
            BorderStyle     =   1  '実線
            Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   2160
            TabIndex        =   83
            Top             =   660
            Width           =   6045
         End
      End
      Begin imNumber6Ctl.imNumber imnPrice_Total 
         Height          =   375
         Left            =   9000
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   5700
         Visible         =   0   'False
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         Calculator      =   "frmYpmf010.frx":1857
         Caption         =   "frmYpmf010.frx":1877
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":18E5
         Keys            =   "frmYpmf010.frx":1903
         Spin            =   "frmYpmf010.frx":1945
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2011496453
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   23
         Left            =   2520
         TabIndex        =   99
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "変更フラグ"
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
      Begin VB.Label Label2 
         Caption         =   "合　計"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         TabIndex        =   63
         Top             =   5760
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   11340
      TabIndex        =   75
      Top             =   0
      Width           =   1635
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
         Left            =   120
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   1935
      Left            =   60
      TabIndex        =   64
      Top             =   1260
      Width           =   12915
      Begin VB.CheckBox chkSoukin 
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1035
      End
      Begin imText6Ctl.imText txtSname 
         Height          =   345
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frmYpmf010.frx":196D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":19DB
         Key             =   "frmYpmf010.frx":19F9
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
      Begin imText6Ctl.imText txtAddres 
         Height          =   405
         Left            =   1560
         TabIndex        =   10
         Top             =   1440
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   714
         Caption         =   "frmYpmf010.frx":1A2D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":1A9B
         Key             =   "frmYpmf010.frx":1AB9
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
         ScrollBars      =   2
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   80
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'なし
         Height          =   495
         Left            =   1500
         TabIndex        =   78
         Top             =   900
         Width           =   2175
         Begin VB.OptionButton optDiv 
            Caption         =   "市外"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1020
            TabIndex        =   8
            Top             =   120
            Width           =   915
         End
         Begin VB.OptionButton optDiv 
            Caption         =   "市内"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   7
            Top             =   120
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdMt070 
         Caption         =   "出品者のマスタ登録(F10)"
         Height          =   375
         Left            =   7620
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   180
         Width           =   2715
      End
      Begin CSComboLib.CSComboBox cboScode 
         Height          =   360
         Left            =   1560
         TabIndex        =   2
         Top             =   180
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
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
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf010.frx":1AFD
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   60
         TabIndex        =   65
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "出品者検索"
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
         Index           =   4
         Left            =   60
         TabIndex        =   66
         Top             =   1440
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "登録番号"
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
         Left            =   9180
         TabIndex        =   67
         Top             =   1440
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
      Begin imText6Ctl.imText txtTel 
         Height          =   360
         Left            =   10680
         TabIndex        =   11
         Top             =   1440
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   635
         Caption         =   "frmYpmf010.frx":1B16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":1B84
         Key             =   "frmYpmf010.frx":1BA2
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
         Text            =   "999999999999999"
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
      Begin CSComboLib.CSComboBox cboScode_Kana 
         Height          =   360
         Left            =   2580
         TabIndex        =   5
         Top             =   180
         Width           =   4875
         _Version        =   262145
         _ExtentX        =   8599
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
         ColWidths       =   "10;20;4"
         Contents        =   "frmYpmf010.frx":1BD6
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   20
         Text            =   "WWWWWWWWWWWWWWWWWWWQ"
         ValueCol        =   2
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   10
         Left            =   60
         TabIndex        =   68
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "出品者名"
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
         Index           =   12
         Left            =   60
         TabIndex        =   77
         Top             =   1020
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "地区区分"
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
         Index           =   22
         Left            =   4200
         TabIndex        =   95
         Top             =   1020
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "送　金"
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
   End
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   3120
      TabIndex        =   55
      Top             =   600
      Width           =   9855
      Begin VB.CommandButton cmdLogin 
         Caption         =   "開催年月日と担当者の変更"
         Height          =   375
         Left            =   6960
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9900
         TabIndex        =   56
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
         Contents        =   "frmYpmf010.frx":1BEF
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   59
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
         TabIndex        =   60
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
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraRecordSelector 
      Height          =   615
      Left            =   7920
      TabIndex        =   46
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
         Picture         =   "frmYpmf010.frx":1C08
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   50
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
         Picture         =   "frmYpmf010.frx":1D52
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   49
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
         Picture         =   "frmYpmf010.frx":1E9C
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   48
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
         Picture         =   "frmYpmf010.frx":1FE6
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   45
      Top             =   9360
      Width           =   12915
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   31
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
         Picture         =   "frmYpmf010.frx":2130
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   11100
         TabIndex        =   34
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
         Picture         =   "frmYpmf010.frx":214C
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   9360
         TabIndex        =   32
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
         Picture         =   "frmYpmf010.frx":22A6
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdPrint 
         Height          =   495
         Left            =   1860
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "ラベル印刷"
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
         rPic.left       =   7
         rPic.top        =   6
         rPic.right      =   21
         rPic.bottom     =   21
         rText.left      =   30
         rText.top       =   8
         rText.right     =   109
         rText.bottom    =   27
         Picture         =   "frmYpmf010.frx":26F8
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdTeita 
         Height          =   495
         Left            =   3660
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   180
         Width           =   2055
         _Version        =   262145
         _ExtentX        =   3625
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "手板用紙印刷"
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
         rPic.left       =   8
         rPic.top        =   6
         rPic.right      =   21
         rPic.bottom     =   21
         rText.left      =   33
         rText.top       =   8
         rText.right     =   133
         rText.bottom    =   27
         Picture         =   "frmYpmf010.frx":2D72
      End
   End
   Begin VB.Frame fraSyori 
      Height          =   615
      Left            =   60
      TabIndex        =   39
      Top             =   0
      Width           =   7815
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
         Left            =   6540
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "確認表印刷"
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   180
         Width           =   1395
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   43
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
   Begin VB.Frame fraKey 
      Height          =   675
      Left            =   60
      TabIndex        =   36
      Top             =   600
      Width           =   3015
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   2340
         Picture         =   "frmYpmf010.frx":33EC
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   180
         Width           =   555
      End
      Begin imText6Ctl.imText txtPnum 
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   180
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   635
         Caption         =   "frmYpmf010.frx":36F6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf010.frx":3764
         Key             =   "frmYpmf010.frx":3782
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
         TabIndex        =   38
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
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   13200
      TabIndex        =   0
      Top             =   180
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf010.frx":37C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf010.frx":3834
      Key             =   "frmYpmf010.frx":3852
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
      Left            =   13380
      TabIndex        =   35
      Top             =   180
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf010.frx":3896
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf010.frx":3904
      Key             =   "frmYpmf010.frx":3922
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
   Begin imText6Ctl.imText imtScode_Kana_Focus1 
      Height          =   135
      Left            =   13200
      TabIndex        =   3
      Top             =   480
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf010.frx":3966
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf010.frx":39D4
      Key             =   "frmYpmf010.frx":39F2
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
   Begin imText6Ctl.imText imtScode_Kana_Focus2 
      Height          =   135
      Left            =   13380
      TabIndex        =   4
      Top             =   480
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf010.frx":3A36
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf010.frx":3AA4
      Key             =   "frmYpmf010.frx":3AC2
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
      Height          =   135
      Left            =   13200
      TabIndex        =   15
      Top             =   3660
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf010.frx":3B06
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf010.frx":3B74
      Key             =   "frmYpmf010.frx":3B92
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
   Begin imText6Ctl.imText imtIcode_Kana_Focus2 
      Height          =   135
      Left            =   13380
      TabIndex        =   16
      Top             =   3660
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf010.frx":3BD6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf010.frx":3C44
      Key             =   "frmYpmf010.frx":3C62
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
Attribute VB_Name = "frmYpmf010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objArPrint As New clsArPrint

Public m_strLastPnum As String                      '最後に登録した受付番号

Const AUTO_CODE = 1                                 'コードの自動採番
Const MAX_ROW = 20                                  '明細の最大行数
Const MAX_COL = 13                                  '明細の列数
Const DETAIL_FORECOLOR1 = &HFF&
Const DETAIL_FORECOLOR2 = &HFF0000

Private Type typDetail
    Div     As Boolean
    Field01 As Variant
    Field02 As Variant
    Field03 As Variant
    Field04 As Variant
    Field05 As Variant
    Field06 As Variant
    Field07 As Variant
    Field08 As Variant
    Field09 As Variant
    Field10 As Variant
    Field11 As Variant
    Field12 As Variant
    Field13 As Variant
End Type
Private m_typDetailCopy As typDetail                '明細のコピー/貼り付け用

Private Sub cboBcode_Click()

    Call cboBcode_Validate(False)

End Sub

Private Sub cboBcode_DropDown()

    Call MakecboBcode(cboBcode)

End Sub

Private Sub cboBcode_GotFocus()
   
    cboBcode.BackColor = FOCUS_STOP_COLOR
    cboBcode.Tag = cboBcode.Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboBcode_LostFocus()
   
    cboBcode.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboBcode_Validate(Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cboBcode_Validate_Err
    
    If Trim(cboBcode.Text) = "" Then
        lblBname.Caption = ""
        Exit Sub
    End If
    If IsNumeric(cboBcode.Text) = False Then
        cboBcode.Text = ""
        lblBname.Caption = ""
        Exit Sub
    End If
    If cboBcode.Tag = cboBcode.Text Then Exit Sub
    
    lblBname.Caption = ""
    
    With adoRecordset1
        '得意先マスタ
        strSQL = "{call sp_MT070;2(" & Trim(cboBcode.Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Fdiv")) Then
                If .Fields("Fdiv") = BUSINESS_DIV_BUYER Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    lblBname.Caption = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If lblBname.Caption = "" Then cboBcode.Text = ""
    
    Exit Sub

cboBcode_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBcode_Validate_Err")

End Sub

Private Sub cboIcode_Kana_Click()
    
    cboIcode_Kana.Tag = "1"
    cboIcode_Kana.BackColor = FOCUS_STOP_COLOR
    Call cboIcode_Kana_Validate(False)

End Sub

Private Sub cboIcode_Kana_DropDown()

    cboIcode_Kana.Tag = "1"
    Call MakecboIcode_Kana(cboIcode_Kana)

End Sub

Private Sub cboIcode_Kana_GotFocus()

    cboIcode_Kana.BackColor = FOCUS_STOP_COLOR
    cboIcode_Kana.Tag = ""
    Call SetImeMode(ActiveControl.hwnd, 9)
    
    cboIcode_Kana.Text = RTrim(cboIcode_Kana.Text)
    cboIcode_Kana.SelStart = 0
    cboIcode_Kana.SelLength = Len(cboIcode_Kana.Text)
    
End Sub

Private Sub cboIcode_Kana_LostFocus()

    cboIcode_Kana.BackColor = FOCUS_NO_COLOR
    cboIcode_Kana.Tag = ""
    
End Sub

Private Sub cboIcode_Kana_Validate(Cancel As Boolean)

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo cboIcode_Kana_Validate_Err
    
    If Trim(cboIcode_Kana.Text) = "" Then Exit Sub
    If Trim(cboIcode_Kana.Tag) = "" Then Exit Sub
    If IsNumeric(cboIcode_Kana.Value) = False Then Exit Sub
    
    txtIname.Text = ""
    cboIcode.Text = cboIcode_Kana.Value
        
    '商品マスタ
    strSQL = "{call sp_MT050;2(" & cboIcode.Text & ")}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        txtIname.Text = IIf(IsNull(adoRecordset1.Fields("Iname")), "", Trim(adoRecordset1.Fields("Iname")))
    End If
    adoRecordset1.Close
    
    Exit Sub

cboIcode_Kana_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboIcode_Kana_Validate_Err")

End Sub

Private Sub cboScode_DropDown()

    Call MakecboScode(cboScode)
    
End Sub

Private Sub cboScode_GotFocus()

    cboScode.BackColor = FOCUS_STOP_COLOR
    cboScode.Tag = cboScode.Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboScode_LostFocus()

    cboScode.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboScode_Kana_Click()
    
    cboScode_Kana.Tag = "1"
    cboScode_Kana.BackColor = FOCUS_STOP_COLOR
    Call cboScode_Kana_Validate(False)
    
End Sub

Private Sub cboScode_Kana_DropDown()

    cboScode_Kana.Tag = "1"
    Call MakecboScode_Kana(cboScode_Kana)

End Sub

Private Sub cboScode_Kana_GotFocus()

    cboScode_Kana.BackColor = FOCUS_STOP_COLOR
    cboScode_Kana.Tag = ""
    Call SetImeMode(ActiveControl.hwnd, 9)

End Sub

Private Sub cboScode_Kana_LostFocus()

    cboScode_Kana.BackColor = FOCUS_NO_COLOR
    cboScode_Kana.Tag = ""

End Sub

Private Sub cboScode_Kana_Validate(Cancel As Boolean)

    On Error GoTo cboScode_Kana_Validate_Err
    
    If Trim(cboScode_Kana.Tag) = "" Then Exit Sub
    If Trim(cboScode_Kana.Tag) = "" Then Exit Sub
    If IsNumeric(cboScode_Kana.Value) = False Then Exit Sub
    
    cboScode.Text = cboScode_Kana.Value
    Call cboScode_Validate(False)
    
    Exit Sub

cboScode_Kana_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboScode_Kana_Validate_Err")

End Sub

Private Sub cboScode_Validate(Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim strBuff As String

    On Error GoTo cboScod_Validate_Err
    
    If Trim(cboScode.Text) = "" Then Exit Sub
    If IsNumeric(cboScode.Text) = False Then
        cboScode.Text = ""
        Exit Sub
    End If
    If cboScode.Tag = cboScode.Text Then Exit Sub
    
    txtSname.Text = ""
    txtAddres.Text = ""
    txtTel.Text = ""
    
    With adoRecordset1
        '得意先マスタ
        strSQL = "{call sp_MT070;2(" & Trim(cboScode.Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If IsNull(.Fields("Fdiv")) = False Then
                If .Fields("Fdiv") = BUSINESS_DIV_EXHIBITION Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    txtSname.Text = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                    txtAddres.Text = IIf(IsNull(.Fields("Addres")), "", Trim(.Fields("Addres")))
                    txtTel.Text = IIf(IsNull(.Fields("Tel")), "", Trim(.Fields("Tel")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If Trim(txtSname.Text) = "" Then cboScode.Text = ""
    
    Exit Sub

cboScod_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboScod_Validate_Err")

End Sub

'目　的　　：
'条　件　　：ｺｰﾄﾞ自動採番クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub chkAutoCode_Click()

    On Error Resume Next

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then
         txtPnum.Text = AutoCodeSet()
'         If txtPnum.Enabled Then txtPnum.SetFocus
        txtPnum.Enabled = False
    ElseIf chkAutoCode.Value = 0 Then
        txtPnum.Enabled = True
        txtPnum.SetFocus
    End If

    Exit Sub

End Sub

Private Sub chkChumon_Click()

    On Error GoTo chkChumon_Click_Err

    If chkChumon.Value = 1 Then
        chkChumon.BackColor = BUTTON_ON
        fraChumon.Visible = True
        imnPrice1.SetFocus
        DoEvents
    Else
        chkChumon.BackColor = BUTTON_OFF
        fraChumon.Visible = False
    End If

    Exit Sub

chkChumon_Click_Err:

    Call MsgBox("注文分クリック時エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "chkChumon_Click_Err")

End Sub

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    On Error Resume Next
    If optSyori(0).Value = True Then
        If chkAutoCode.Value = 1 Then txtPnum.Text = AutoCodeSet()
    End If
    txtPnum.SetFocus

End Sub

Private Sub cmdClear_Dst_Click()

    Call FieldsClear(3)
    Call ListViewGetMaxRow
    cboIcode_Kana.SetFocus

End Sub

Private Sub cmdCopy_Click()

    On Error GoTo cmdCopy_Click_Err

    Call ListViewGetMaxRow
    '明細クリア
    chkIdiv.Value = 0
    chkSdiv.Value = 0
    chkBdiv.Value = 0
    imnBnum.Value = 0
    imnSnum.Value = 0
    
    m_typDetailCopy.Div = True
    m_typDetailCopy.Field01 = Trim(cboIcode.Text)
    m_typDetailCopy.Field02 = txtIname.Text
    m_typDetailCopy.Field03 = Format(imnQty.Value, "#,##0")
    m_typDetailCopy.Field04 = chkIdiv.Value
    m_typDetailCopy.Field05 = Format(imnPrice1.Value, "#,##0")
    m_typDetailCopy.Field06 = Format(imnPrice1.Value, "#,##0")
    m_typDetailCopy.Field07 = Format(imnPrice.Value, "#,##0")
    m_typDetailCopy.Field08 = Trim(cboBcode.Text)
    m_typDetailCopy.Field09 = lblBname.Caption
    m_typDetailCopy.Field10 = chkSdiv.Value
    m_typDetailCopy.Field11 = chkBdiv.Value
    m_typDetailCopy.Field12 = imnBnum.Value
    m_typDetailCopy.Field13 = imnSnum.Value
    
    Exit Sub

cmdCopy_Click_Err:

    Call MsgBox("明細コピークリック時エラー！！" _
            & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdCopy_Click_Err")

End Sub

'目　的　　：
'条　件　　：レコード移動クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub cmdDataMove_Click(Index As Integer)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cmdDataMove_Click_Err

    Screen.MousePointer = vbHourglass

    With adoRecordset1
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            Select Case Index
                Case 0:
                    .MoveFirst
                Case 1:
                    If Trim(txtPnum.Text) = "" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    .Find "Pnum = " & Trim(txtPnum.Text)
                    If Not .EOF Then
                        .MovePrevious
                        If .EOF Or .BOF Then .MoveFirst
                    Else
                        .MoveFirst
                    End If
                Case 2:
                    If Trim(txtPnum.Text) = "" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    .Find "Pnum = " & Trim(txtPnum.Text)
                    If Not .EOF Then
                        .MoveNext
                        If .EOF Or .BOF Then .MoveLast
                    Else
                        .MoveLast
                    End If
                Case 3:
                    .MoveLast
                Case Else
                    Exit Sub
            End Select
            Call FieldsClear(0)
            If FieldsSet(True, adoRecordset1) = False Then Exit Sub
        End If
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
cmdDataMove_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("レコード移動クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdDataMove_Click_Err")
    
End Sub

Private Sub cmdDel_Click()

    If ListViewDelItem() = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    cboIcode_Kana.SetFocus

End Sub

Private Sub cmdDetail_Click()

    fraDetail.Visible = Not fraDetail.Visible

End Sub

Private Sub cmdEdit_Click()

    If DoValidationChecks_Dst() = False Then Exit Sub
    If ListViewSetItem(imnNo.Value, 0) = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    
    cboIcode_Kana.SetFocus

End Sub

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error Resume Next

    If MsgBox("実行しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    If optSyori(0).Value = True Or optSyori(1).Value = True Then
        '入力チェック
        If DoValidationChecks() = False Then Exit Sub
        If DataUpdate() = False Then Exit Sub
        m_strLastPnum = txtPnum.Text
'        If optSyori(0).Value = True Then
'            If MsgBox("確認表を印刷しますか？", vbYesNo + vbQuestion, "") = vbYes Then
'                frmPrintDialog.m_blnAutoPrint = True
'                frmPrintDialog.Show vbModal
'            End If
'        End If
           
        '2005/08/12 削除
        '出品伝票
'        Call cmdPrint_Click
           
    ElseIf optSyori(2).Value = True Then
        '入力チェック
        If DoValidationChecks() = False Then Exit Sub
        If DataDelete() = False Then Exit Sub
    End If
    
    'フィールドクリア
    Call FieldsClear(0)

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then
        txtPnum.Text = AutoCodeSet
        txtSname.SetFocus
    Else
        txtPnum.Enabled = True
        txtPnum.SetFocus
    End If
    
    DoEvents
    
End Sub

'目　的　　：
'条　件　　：終了クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdInsert_Click()

    Call SetImeMode(ActiveControl.hwnd, 2)
    If DoValidationChecks_Dst() = False Then Exit Sub
    If ListViewInsItem() = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    cboIcode_Kana.SetFocus

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
        If optSyori(0).Value = True Then
            Call optSyori_Click(0)
        Else
            optSyori(0).Value = True
        End If
    End If
    Unload frmLogin
    m_strLastPnum = ""

    Exit Sub

cmdLogin_Click_Err:

    Call MsgBox("開催年月日と担当者の変更クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdLogin_Click_Err")

End Sub

'目　的　　：
'条　件　　：植木のマスタ登録クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub cmdMt050_Click()

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim lngCode As Long
    Dim blnAddNew As Boolean

    On Error GoTo cmdMt050_Click_Err

    If MsgBox("植木をマスタ登録しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    If Trim(txtIname.Text) = "" Then
        DoEvents
        Call MsgBox("植木名を入力してください。", vbOKOnly + vbCritical, "")
        txtIname.SetFocus
        DoEvents
        Exit Sub
    End If
    If Trim(cboIcode_Kana.Text) = "" Then
        DoEvents
        Call MsgBox("カナを入力してください。", vbOKOnly + vbCritical, "")
        cboIcode_Kana.SetFocus
        DoEvents
        Exit Sub
    End If
    
    blnAddNew = True
    With adoRecordset1
        '商品マスタ
        strSQL = "SELECT * FROM MT050" & _
                 " WHERE Iname = '" & Trim(txtIname.Text) & "'" & _
                 " ORDER BY Icode"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            '更新
            blnAddNew = False
            lngCode = .Fields("Icode")
            .Fields("Iname") = txtIname.Text
            .Fields("Ikana") = Global_LeftB_Ansi(cboIcode_Kana.Text, 40)
            .Update
        Else
            .Close
            blnAddNew = True
            '商品マスタ
            strSQL = "SELECT * FROM MT050" & _
                     " ORDER BY Icode DESC"
            .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If Not .EOF Then
                lngCode = CLng(.Fields("Icode")) + 1
            Else
                lngCode = 1
            End If
            .AddNew
            .Fields("Icode") = lngCode
            .Fields("Iname") = txtIname.Text
            .Fields("Ikana") = Global_LeftB_Ansi(cboIcode_Kana.Text, 40)
            .Fields("Idiv") = 1
            .Update
            .Close
        End If
    End With
        
    cboIcode.Text = lngCode
        
    txtIname.SetFocus
    DoEvents
    If blnAddNew Then
        Call MsgBox("新規登録しました。", vbOKOnly + vbInformation, "")
    Else
        Call MsgBox("更新登録しました。", vbOKOnly + vbInformation, "")
    End If
        
    Exit Sub

cmdMt050_Click_Err:

    Call MsgBox("植木のマスタ登録クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdMt050_Click_Err")

End Sub

'目　的　　：
'条　件　　：出品者のマスタ登録クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub cmdMt070_Click()

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim lngCode As Long
    Dim blnAddNew As Boolean

    On Error GoTo cmdMt070_Click_Err

    If MsgBox("出品者をマスタ登録しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    If Trim(txtSname.Text) = "" Then
        DoEvents
        Call MsgBox("出品者名を入力してください。", vbOKOnly + vbCritical, "")
        txtSname.SetFocus
        DoEvents
        Exit Sub
    End If
    
    blnAddNew = True
    With adoRecordset1
        If Trim(cboScode.Text) = "" Or IsNumeric(cboScode.Text) = False Then
            '得意先マスタ
            strSQL = "SELECT * FROM MT070" & _
                     " WHERE Bname = '" & txtSname.Text & "'" & _
                     " ORDER BY Bcode"
        Else
            '得意先マスタ
            strSQL = "SELECT * FROM MT070" & _
                     " WHERE Bcode = " & cboScode.Text & _
                     " ORDER BY Bcode"
        End If
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            '更新
            blnAddNew = False
            lngCode = .Fields("Bcode")
            .Fields("Bcode") = lngCode
            .Fields("Bname") = txtSname.Text
            .Fields("Bkana") = Global_LeftB_Ansi(cboScode_Kana.Text, 20)
            .Fields("Addres") = txtAddres.Text
            .Fields("Tel") = txtTel.Text
            .Update
        Else
            .Close
            blnAddNew = True
            '得意先マスタ
            strSQL = "SELECT * FROM MT070" & _
                     " ORDER BY Bcode DESC"
            .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If Not .EOF Then
                lngCode = CLng(.Fields("Bcode")) + 1
            Else
                lngCode = 1
            End If
            .AddNew
            .Fields("Bcode") = lngCode
            .Fields("Bname") = txtSname.Text
            .Fields("Bkana") = Global_LeftB_Ansi(cboScode_Kana.Text, 20)
            .Fields("Addres") = txtAddres.Text
            .Fields("Tel") = txtTel.Text
            .Fields("Rname") = txtSname.Text
            .Fields("Rdiv") = RECEIPT_ON
            .Fields("Fdiv") = BUSINESS_DIV_EXHIBITION
            .Update
            .Close
        End If
    End With
        
    cboScode.Text = lngCode
        
    txtSname.SetFocus
    DoEvents
    If blnAddNew Then
        Call MsgBox("新規登録しました。", vbOKOnly + vbInformation, "")
    Else
        Call MsgBox("更新登録しました。", vbOKOnly + vbInformation, "")
    End If
        
    Exit Sub

cmdMt070_Click_Err:

    Call MsgBox("出品者のマスタ登録クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdMt070_Click_Err")

End Sub

Private Sub cmdPast_Click()

    On Error GoTo cmdPast_Click_Err

    If m_typDetailCopy.Div = False Then Exit Sub

    Call ListViewGetMaxRow
    cboIcode.Text = m_typDetailCopy.Field01
    txtIname.Text = m_typDetailCopy.Field02
    imnQty.Value = m_typDetailCopy.Field03
    chkIdiv.Value = m_typDetailCopy.Field04
    imnPrice1.Value = m_typDetailCopy.Field05
    imnPrice1.Value = m_typDetailCopy.Field06
    imnPrice.Value = m_typDetailCopy.Field07
    cboBcode.Text = m_typDetailCopy.Field08
    lblBname.Caption = m_typDetailCopy.Field09
'    chkSdiv.Value = m_typDetailCopy.Field10
'    chkBdiv.Value = m_typDetailCopy.Field11
'    imnBnum.Value = m_typDetailCopy.Field12
'    imnSnum.Value = m_typDetailCopy.Field13
    chkSdiv.Value = 0
    chkBdiv.Value = 0
    imnBnum.Value = 0
    imnSnum.Value = 0
    
    If DoValidationChecks_Dst() = False Then Exit Sub
    If ListViewSetItem(imnNo.Value, 0) = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    
    Exit Sub

cmdPast_Click_Err:

    Call MsgBox("明細貼付クリック時エラー！！" _
            & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdPast_Click_Err")

End Sub

'目　的　　：
'条　件　　：伝票発行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００３／０４／１１
'更新履歴　：
'
Private Sub cmdPrint_Click()
    
    If Trim(txtPnum.Text) = "" Then Exit Sub
    If optSyori(0).Value = False And optSyori(1).Value = False Then Exit Sub
    
    If optSyori(0).Value = True Then
        frmPrintDialog2.m_intMode = 0
    ElseIf optSyori(1).Value = True Then
        frmPrintDialog2.m_intMode = 1
    End If
    frmPrintDialog2.m_intPnum = txtPnum.Text
    Beep
    frmPrintDialog2.Show vbModal
    
End Sub

'目　的　　：
'条　件　　：検索クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub cmdSearch_Click()

    frmSearch.Show vbModal

End Sub

Private Sub cmdTeita_Click()
    
    Dim objRpt As New rptYpmf011
    
    On Error GoTo cmdTeita_Click_Err
    
    If Trim(txtPnum.Text) = "" Then
        MsgBox "受付番号を入力してください。", vbOKOnly + vbInformation, ""
        Exit Sub
    End If
    
    If MakeWork() = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "競売結果確認表"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_YPMF012"
        .Caption = "競売結果確認表"
        If .PrintActiveReport(0) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End With
    
    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
cmdTeita_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("実行クリックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdTeita_Click_Err")
    
End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
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
            If cmdMt070.Enabled = False Then Exit Sub
            cmdMt070.SetFocus
            DoEvents
            Call cmdMt070_Click
        Case vbKeyF11
            If cmdMt050.Enabled = False Then Exit Sub
            cmdMt050.SetFocus
            DoEvents
            Call cmdMt050_Click
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
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "出品者受付入力"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    chkAutoCode.Value = AUTO_CODE

    '処理ボタン
    optSyori(0).Value = True
    optSyori(1).Value = False
    optSyori(2).Value = False
    optSyori(3).Value = False
    optSyori(4).Value = False

'    chkAutoCode.Value = AUTO_CODE
'    If chkAutoCode.Value = 1 Then txtPnum.Text = AutoCodeSet()
    m_strLastPnum = ""
    
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
'作成年月日：２００２／０６／１３
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
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    Dim strSQL As String

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        txtPnum.Text = ""
        txtPnum.Tag = ""
        If optSyori(0).Value = True Then
            txtPnum.Enabled = False
        End If
        
        'ヘッダー
        cboScode_Kana.Text = ""
        cboScode_Kana.Tag = ""
        txtSname.Text = ""
        cboScode.Text = ""
        cboScode.Tag = ""
        txtAddres.Text = ""
        txtTel.Text = ""
        optDiv(0).Value = True
        chkSoukin.Value = 0
        '明細
        imnNo.Value = 1
        cboIcode.Text = ""
        cboIcode.Tag = ""
        cboIcode_Kana.Text = ""
        txtIname.Text = ""
        imnQty.Value = 1            '数量の初期値「1」
        chkIdiv.Value = 0
        imnPrice1.Value = 0
        imnPrice2.Value = 0
        imnPrice.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
        'フッター
        lsvMeisai.ListItems.Clear
        imnQty_Total.Value = 0
        imnPrice_Total.Value = 0
        chkChumon.Value = 0
        fraChumon.Visible = False
        
        m_typDetailCopy.Div = False
    ElseIf intKubun = 1 Then
        txtPnum.Text = ""
        txtPnum.Tag = ""
    ElseIf intKubun = 2 Then
        'ヘッダー
        cboScode_Kana.Text = ""
        cboScode_Kana.Tag = ""
        txtSname.Text = ""
        cboScode.Text = ""
        cboScode.Tag = ""
        txtAddres.Text = ""
        txtTel.Text = ""
        optDiv(0).Value = True
        chkSoukin.Value = 0
        '明細
        imnNo.Value = 1
        cboIcode.Text = ""
        cboIcode.Tag = ""
        cboIcode_Kana.Text = ""
        txtIname.Text = ""
        imnQty.Value = 1            '数量の初期値「1」
        imnPrice1.Value = 0
        imnPrice2.Value = 0
        imnPrice.Value = 0
        chkIdiv.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
        'フッター
        lsvMeisai.ListItems.Clear
        imnQty_Total.Value = 0
        imnPrice_Total.Value = 0
        chkChumon.Value = 0
        fraChumon.Visible = False
        
        m_typDetailCopy.Div = False
    ElseIf intKubun = 3 Then
        '明細
        cboIcode.Text = ""
        cboIcode_Kana.Text = ""
        txtIname.Text = ""
        imnQty.Value = 1            '数量の初期値「1」
        chkIdiv.Value = 0
        imnPrice1.Value = 0
        imnPrice2.Value = 0
        imnPrice.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
    End If
        
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF011"
    g_clsAdoAccess.Connection.Execute strSQL
        
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

Private Sub imnBnum_GotFocus()
    
    imnBnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBnum_LostFocus()
    
    imnBnum.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnPrice1_GotFocus()
    
    imnPrice1.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnPrice1_LostFocus()
    
    imnPrice1.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnPrice1_Validate(Cancel As Boolean)

    If Calc_Price() = False Then Cancel = True
    
End Sub

Private Sub imnPrice2_GotFocus()
    
    imnPrice2.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnPrice2_LostFocus()
    
    imnPrice2.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnPrice2_Validate(Cancel As Boolean)

    If imnPrice2.Value > imnPrice1.Value Then
        If MsgBox("仕入単価のほうが高いですが、よろしいですか？", vbYesNo + vbInformation, "") = vbNo Then
            Cancel = True
        End If
    End If

    If Calc_Price() = False Then Cancel = True
    
End Sub

Private Sub imnQty_GotFocus()
    
    imnQty.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnQty_LostFocus()
    
    imnQty.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnQty_Validate(Cancel As Boolean)

    If Calc_Price() = False Then Cancel = True

End Sub

Private Sub imnSnum_GotFocus()
    
    imnSnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnSnum_LostFocus()
    
    imnSnum.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imtIcode_Kana_Focus1_GotFocus()

'    On Error Resume Next

'    If Trim(txtIname.Text) = "" Then
        txtIname.SetFocus
'    Else
'        imnQty.SetFocus
'    End If

End Sub

Private Sub imtIcode_Kana_Focus2_GotFocus()

    cboIcode_Kana.SetFocus

End Sub

Private Sub imtScode_Kana_Focus1_GotFocus()

    If Trim(txtSname.Text) = "" Then
        txtSname.SetFocus
    Else
        If optSyori(0).Value = True Or optSyori(1).Value = True Then
            cboIcode_Kana.SetFocus
        Else
            cmdExecute.SetFocus
        End If
    End If

End Sub

Private Sub imtScode_Kana_Focus2_GotFocus()

    cboScode.SetFocus

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    If optSyori(0).Value = True Then
        txtSname.SetFocus
    Else
        If txtPnum.Enabled = True Then
            txtPnum.SetFocus
        End If
    End If

End Sub

Private Sub lsvMeisai_Click()

    On Error Resume Next

    '行が選択されているか？
    If lsvMeisai.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    '明細表示
    Call ListViewGetItem
    
    cboIcode_Kana.SetFocus

End Sub

'目　的　　：
'条　件　　：処理区分ボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub optSyori_Click(Index As Integer)

    Dim intIndex1 As Integer
    Dim adoRecordset1 As New ADODB.Recordset

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
            Call FieldsControl(2, True)
            Call FieldsControl(3, True)
            If chkAutoCode.Value = 1 Then
                txtPnum.Text = AutoCodeSet
            Else
                txtPnum.Enabled = True
            End If
        Case 1: '変更
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
            Call FieldsControl(3, True)
            txtPnum.Enabled = True
        Case 2: '削除
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
            Call FieldsControl(3, False)
            txtPnum.Enabled = True
        Case 3: '印刷
            Call FieldsControl(0, True)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
            Call FieldsControl(3, False)
            frmPrintDialog.m_blnAutoPrint = False
            frmPrintDialog.Show vbModal
        Case 4: '外部出力
    End Select

    On Error Resume Next
    If Index = 0 And chkAutoCode.Value = 1 Then
        txtSname.SetFocus
    Else
        txtPnum.SetFocus
    End If
    DoEvents
    
    Exit Sub

optSyori_Click_Err:

    Call MsgBox("処理区分クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")

End Sub

Private Sub txtAddres_GotFocus()

    txtAddres.BackColor = FOCUS_STOP_COLOR
    Call SetImeMode(ActiveControl.hwnd, 4)

End Sub

Private Sub txtAddres_LostFocus()

    txtAddres.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtIname_GotFocus()

    txtIname.BackColor = FOCUS_STOP_COLOR
    Call SetImeMode(ActiveControl.hwnd, 4)

End Sub

Private Sub txtIname_LostFocus()

    txtIname.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtPnum_Change()

    If Trim(txtPnum.Text) = "" Then Exit Sub

    If txtPnum.Tag <> txtPnum.Text Then
        If optSyori(0).Value Or optSyori(1).Value Then
            fraMeisai.Enabled = True
            DoEvents
        End If
    End If

End Sub

Private Sub txtPnum_Click()

    Call cboScode_Validate(False)
    
End Sub

Private Sub txtPnum_GotFocus()

    txtPnum.Tag = txtPnum.Text
    txtPnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtPnum_LostFocus()

    txtPnum.Tag = ""
    txtPnum.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtPnum_Validate(Cancel As Boolean)

    If Trim(txtPnum.Text) = "" Then Exit Sub
    If txtPnum.Tag = txtPnum.Text Then Exit Sub

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

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "開催年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtPnum.Text) = "" Then
        strErrMsg = "受付番号を入力してください。"
        txtPnum.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtSname.Text) = "" Then
        strErrMsg = "出品者名を入力してください。"
        txtSname.SetFocus
        GoTo ErrorTrap:
    End If
    If lsvMeisai.ListItems.Count <= 0 Then
        strErrMsg = "明細を入力してください。"
        GoTo ErrorTrap:
    End If
    
    '新規で自動採番以外はチェック
    If optSyori(0).Value = True And chkAutoCode.Value = 0 Then
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & txtPnum.Text
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoRecordset1.EOF = False Then
            strErrMsg = "受付番号:" & txtPnum.Text & "は既に登録されています。"
            GoTo ErrorTrap:
        End If
        adoRecordset1.Close
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
'引　数　　：intKbn 0:キー部 1:レコード移動 2:明細　3:ヘッダー
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
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
        Case 3:
            fraHeader.Enabled = blnEnabled
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
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Public Function FieldsSet(blnVisible As Boolean, Optional adoRecordsetArg As Variant) As Boolean
    
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim strSQL As String
    Dim itmX As ListItem
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim strBuff As String
    Dim varColor As Variant

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    If IsMissing(adoRecordsetArg) = False Then
        Set adoRecordset1 = adoRecordsetArg
    Else
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & txtPnum.Text
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
        txtPnum.Text = .Fields("Pnum")
        cboScode.Text = IIf(IsNull(.Fields("Scode")), "", Trim(.Fields("Scode")))
        txtSname.Text = IIf(IsNull(.Fields("Sname")), "", Trim(.Fields("Sname")))
        txtAddres.Text = IIf(IsNull(.Fields("Addres")), "", Trim(.Fields("Addres")))
        txtTel.Text = IIf(IsNull(.Fields("Tel")), "", Trim(.Fields("Tel")))
        If Not IsNull(.Fields("Div")) Then
            If .Fields("Div") = TIKU_DIV_OFF Then
                optDiv(1).Value = True
            ElseIf .Fields("Div") = TIKU_DIV_ON Then
                optDiv(0).Value = True
            End If
        End If
        chkSoukin.Value = IIf(IsNull(.Fields("Soukin")), 0, .Fields("Soukin"))
        .Close
    End With
    
    With adoRecordset2
        intIndex1 = 1
        lsvMeisai.ListItems.Clear
        
        strSQL = "SELECT * FROM DT011" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & txtPnum.Text & _
                 " ORDER BY Odate,Pnum,Line"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Set itmX = lsvMeisai.ListItems.Add(, , intIndex1, 0)
            itmX.SubItems(1) = IIf(IsNull(.Fields("Icode")), "", .Fields("Icode"))
            itmX.SubItems(2) = IIf(IsNull(.Fields("Iname")), "", Trim(.Fields("Iname")))
            itmX.SubItems(3) = IIf(IsNull(.Fields("Qty")), 0, Format(.Fields("Qty"), "#,##0"))
            itmX.SubItems(4) = IIf(IsNull(.Fields("Idiv")), 0, .Fields("Idiv"))
            itmX.SubItems(5) = IIf(IsNull(.Fields("Price1")), "", Format(.Fields("Price1"), "#,##0"))
            itmX.SubItems(6) = IIf(IsNull(.Fields("Price2")), "", Format(.Fields("Price2"), "#,##0"))
            itmX.SubItems(7) = IIf(IsNull(.Fields("Price")), "", Format(.Fields("Price"), "#,##0"))
            itmX.SubItems(8) = IIf(IsNull(.Fields("Bcode")), "", .Fields("Bcode"))
            If Not IsNull(.Fields("Bcode")) Then
                itmX.SubItems(9) = IIf(IsNull(.Fields("Bcode")), "", Global_Get_Bname(g_clsAdoSQL, .Fields("Bcode"), lblOdate.Caption, strBuff))
            Else
                itmX.SubItems(9) = ""
            End If
            itmX.SubItems(10) = IIf(IsNull(.Fields("Sdiv")), 0, .Fields("Sdiv"))
            itmX.SubItems(11) = IIf(IsNull(.Fields("Bdiv")), 0, .Fields("Bdiv"))
            itmX.SubItems(12) = IIf(IsNull(.Fields("Bnum")), 0, .Fields("Bnum"))
            itmX.SubItems(13) = IIf(IsNull(.Fields("Snum")), 0, .Fields("Snum"))
            itmX.SubItems(14) = 0
            
            '入力済みの場合は、前景色を変える
            If itmX.SubItems(4) = INPUT_ON Or itmX.SubItems(7) <> "0" Then
                varColor = DETAIL_FORECOLOR2
            Else
                varColor = DETAIL_FORECOLOR1
            End If
            itmX.ForeColor = varColor
            For intIndex2 = 1 To MAX_COL
                itmX.ListSubItems(intIndex2).ForeColor = varColor
            Next intIndex2

            intIndex1 = intIndex1 + 1
            .MoveNext
        Loop
        .Close
        
        Call ListViewGetMaxRow
    End With

    Call Calc_Total
    
    Set adoRecordset1 = Nothing
    Set adoRecordset2 = Nothing
    
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
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim wkRecordset1 As New ADODB.Recordset
    Dim intIndex1 As Integer

    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
    If optSyori(0).Value = True Then
        If chkAutoCode.Value = 1 Then
            txtPnum.Text = AutoCodeSet()
        End If
    ElseIf optSyori(1).Value = True Then
'        '出品者精算データ
'        strSQL = "SELECT * FROM DT040" & _
'                 " WHERE Odate = '" & Trim(lblOdate.Caption) & "'" & _
'                 " AND Pnum = " & txtPnum.Text & _
'                 " ORDER BY Odate,Pnum,Num DESC"
'        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
'        If adoRecordset1.EOF = False Then
'            g_clsAdoSQL.Connection.RollbackTrans
'            DataUpdate = False
'            Screen.MousePointer = vbDefault
'            Call MsgBox("既に精算されているため変更できません。", vbOKOnly + vbCritical, "")
'            Exit Function
'        End If
'        adoRecordset1.Close
    End If
 
    strSQL = "DELETE FROM DT011" & _
             " WHERE Odate = '" & lblOdate.Caption & "'" & _
             " AND Pnum = " & txtPnum.Text
    g_clsAdoSQL.Connection.Execute strSQL
    
    strSQL = "DELETE FROM DT010" & _
             " WHERE Odate = '" & lblOdate.Caption & "'" & _
             " AND Pnum = " & txtPnum.Text
    g_clsAdoSQL.Connection.Execute strSQL
 
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF011"
    g_clsAdoAccess.Connection.Execute strSQL
    
    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF011"
    wkRecordset1.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
 
    With adoRecordset1
        strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & txtPnum.Text
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        .Fields("Odate") = lblOdate.Caption
        .Fields("Pnum") = txtPnum.Text
        If IsNumeric(cboScode.Text) = True Then .Fields("Scode") = cboScode.Text
        .Fields("Sname") = txtSname.Text
        .Fields("Addres") = txtAddres.Text
        .Fields("Tel") = txtTel.Text
        If optDiv(0).Value = True Then
            .Fields("Div") = TIKU_DIV_ON
        ElseIf optDiv(1).Value = True Then
            .Fields("Div") = TIKU_DIV_OFF
        End If
        .Fields("Soukin") = chkSoukin.Value
        .Update
        .Close
    End With
    
    With adoRecordset2
        strSQL = "SELECT * FROM DT011"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            .AddNew
            .Fields("Odate") = lblOdate.Caption
            .Fields("Pnum") = txtPnum.Text
            .Fields("Line") = intIndex1
            .Fields("Icode") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)), lsvMeisai.ListItems(intIndex1).SubItems(1), Null)
            .Fields("Iname") = lsvMeisai.ListItems(intIndex1).SubItems(2)
            .Fields("Qty") = lsvMeisai.ListItems(intIndex1).SubItems(3)
            .Fields("Price1") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(5)), lsvMeisai.ListItems(intIndex1).SubItems(5), Null)
            .Fields("Price2") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(6)), lsvMeisai.ListItems(intIndex1).SubItems(6), Null)
            .Fields("Price") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(7)), lsvMeisai.ListItems(intIndex1).SubItems(7), Null)
            .Fields("Bcode") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(8)), lsvMeisai.ListItems(intIndex1).SubItems(8), Null)
            .Fields("Sdiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(10)), lsvMeisai.ListItems(intIndex1).SubItems(10), 0)
            .Fields("Bdiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(11)), lsvMeisai.ListItems(intIndex1).SubItems(11), 0)
            .Fields("Bnum") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(12)), lsvMeisai.ListItems(intIndex1).SubItems(12), 0)
            .Fields("Snum") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(13)), lsvMeisai.ListItems(intIndex1).SubItems(13), 0)
            .Fields("Idiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(4)), lsvMeisai.ListItems(intIndex1).SubItems(4), 0)
            .Update
        
            If lsvMeisai.ListItems(intIndex1).SubItems(14) = "1" Then
                wkRecordset1.AddNew
                wkRecordset1.Fields("Odate") = lblOdate.Caption
                wkRecordset1.Fields("Pnum") = txtPnum.Text
                wkRecordset1.Fields("Line") = intIndex1
                wkRecordset1.Update
            End If
        Next
        .Close
    End With
    
    g_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    Set wkRecordset1 = Nothing
    Set adoRecordset1 = Nothing
    Set adoRecordset2 = Nothing
    
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
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function DataDelete() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    
    On Error GoTo DataDelete_Err
    
    Screen.MousePointer = vbHourglass
    
'    '出品者精算データ
'    strSQL = "SELECT * FROM DT040" & _
'             " WHERE Odate = '" & Trim(lblOdate.Caption) & "'" & _
'             " AND Pnum = " & txtPnum.Text & _
'             " ORDER BY Odate,Pnum,Num DESC"
'    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
'    If adoRecordset1.EOF = False Then
'        DataDelete = False
'        Screen.MousePointer = vbDefault
'        Call MsgBox("既に精算されているため削除できません。", vbOKOnly + vbCritical, "")
'        Exit Function
'    End If
'    adoRecordset1.Close
    
    With g_clsAdoSQL.Connection
        .BeginTrans
        
        strSQL = "DELETE FROM DT011" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & txtPnum.Text
        .Execute strSQL
    
        strSQL = "DELETE FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Pnum = " & txtPnum.Text
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

'目　的　　：コードの自動採番
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function AutoCodeSet() As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo AutoCodeSet_Err
    
    AutoCodeSet = ""
    
    Screen.MousePointer = vbHourglass
    
    With adoRecordset1
        strSQL = "SELECT Pnum FROM DT010" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " ORDER BY Odate,Pnum"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If .EOF Or .BOF Then
            AutoCodeSet = 1
            adoRecordset1.Close
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        .MoveLast
        If CLng(.Fields("Pnum")) < 9999 Then
            AutoCodeSet = CLng(.Fields("Pnum")) + 1
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Function

AutoCodeSet_Err:

    AutoCodeSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("コードの自動採番エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "AutoCodeSet_Err")

End Function

'目　的　　：リストビューへのデータ登録
'条　件　　：
'結　果　　：
'引　数　　：intFlg(0:追加・更新 1:挿入)
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function ListViewSetItem(intPostion As Integer, intFlg As Integer) As Boolean

    Dim itmX As ListItem
    Dim intIndex1 As Integer
    Dim varColor As Variant

    On Error GoTo ListViewSetItem_Err
    
    ListViewSetItem = False
    
    'リストビューのデータ検索（行番号が一致するデータがあったら削除）
    Set itmX = lsvMeisai.FindItem(intPostion, , , 0)
    If Not (itmX Is Nothing) Then
        If intFlg = 0 Then
            'データ削除
            lsvMeisai.ListItems.Remove itmX.Index
        End If
        'データを追加
        Set itmX = lsvMeisai.ListItems.Add(intPostion, , intPostion, 0)
    Else
        'データを追加
        Set itmX = lsvMeisai.ListItems.Add(, , intPostion, 0)
    End If
    itmX.SubItems(1) = Trim(cboIcode.Text)
    itmX.SubItems(2) = txtIname.Text
    itmX.SubItems(3) = Format(imnQty.Value, "#,##0")
    itmX.SubItems(4) = chkIdiv.Value
    itmX.SubItems(5) = Format(imnPrice1.Value, "#,##0")
    itmX.SubItems(6) = Format(imnPrice2.Value, "#,##0")
    itmX.SubItems(7) = Format(imnPrice.Value, "#,##0")
    itmX.SubItems(8) = Trim(cboBcode.Text)
    itmX.SubItems(9) = lblBname.Caption
    itmX.SubItems(10) = chkSdiv.Value
    itmX.SubItems(11) = chkBdiv.Value
    itmX.SubItems(12) = imnBnum.Value
    itmX.SubItems(13) = imnSnum.Value
    itmX.SubItems(14) = 1
    
'    '入力済みの場合は、前景色を変える
'    If itmX.SubItems(4) = INPUT_ON Or itmX.SubItems(7) <> "0" Then
'        varColor = DETAIL_FORECOLOR2
'    Else
'        varColor = DETAIL_FORECOLOR1
'    End If
'    itmX.ForeColor = varColor
'    For intIndex1 = 1 To MAX_COL
'        itmX.ListSubItems(intIndex1).ForeColor = varColor
'    Next intIndex1
    
    'リストビューをスクロールして、検出された ListItem を表示
    lsvMeisai.ListItems(lsvMeisai.ListItems.Count).EnsureVisible
    
    '行番号取得
    Call ListViewGetMaxRow
    
    ListViewSetItem = True
    
    Exit Function

ListViewSetItem_Err:

    Call MsgBox("リストビューへのデータ登録エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewSetItem_Err")

End Function

'目　的　　：リストビューからのデータ表示
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub ListViewGetItem()

    On Error GoTo ListViewGetItem_Err
    
    imnNo.Value = lsvMeisai.SelectedItem.Text
    cboIcode.Text = Trim(lsvMeisai.SelectedItem.SubItems(1))
    txtIname.Text = Trim(lsvMeisai.SelectedItem.SubItems(2))
    imnQty.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(3)) <> "", lsvMeisai.SelectedItem.SubItems(3), 0)
    chkIdiv.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(4)), lsvMeisai.SelectedItem.SubItems(4), 0)
    imnPrice1.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(5)), lsvMeisai.SelectedItem.SubItems(5), 0)
    imnPrice2.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(6)), lsvMeisai.SelectedItem.SubItems(6), 0)
    imnPrice.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(7)), lsvMeisai.SelectedItem.SubItems(7), 0)
    cboBcode.Text = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(8)), lsvMeisai.SelectedItem.SubItems(8), "")
    lblBname.Caption = Trim(lsvMeisai.SelectedItem.SubItems(9))
    chkSdiv.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(10)), lsvMeisai.SelectedItem.SubItems(10), 0)
    chkBdiv.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(11)), lsvMeisai.SelectedItem.SubItems(11), 0)
    imnBnum.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(12)), lsvMeisai.SelectedItem.SubItems(12), 0)
    imnSnum.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(13)), lsvMeisai.SelectedItem.SubItems(13), 0)
    chkUpdflg.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(14)), lsvMeisai.SelectedItem.SubItems(14), 0)
            
    '金額がゼロ以外は注文分とみなす
    If imnPrice.Value <> 0 Then
        chkChumon.Value = 1
    End If
        
    Exit Sub
    
ListViewGetItem_Err:

   Call MsgBox("リストビューからデータ取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewGetItem_Err")

End Sub

'目　的　　：リストビューからのデータ削除
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function ListViewDelItem() As Boolean

    Dim itmX As ListItem
    Dim intPostion As Integer

    On Error GoTo ListViewDelItem_Err

    ListViewDelItem = False

    If MsgBox("明細を削除しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Function
    
    '削除行の取得
    intPostion = imnNo.Value
    
    'リストビューのデータ検索（行番号が一致するデータがあったら削除）
    Set itmX = lsvMeisai.FindItem(intPostion, , , 0)
    If Not (itmX Is Nothing) Then
        'データ削除
        lsvMeisai.ListItems.Remove itmX.Index
        '行番号振り直し
        Call ListViewRefresh
    End If

    '行番号取得
    Call ListViewGetMaxRow

    ListViewDelItem = True

    Exit Function

ListViewDelItem_Err:

    Call MsgBox("リストビューからデータ削除エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewDelItem_Err")

End Function

Private Sub txtSname_GotFocus()

    txtSname.BackColor = FOCUS_STOP_COLOR
    Call SetImeMode(ActiveControl.hwnd, 4)

End Sub

Private Sub txtSname_LostFocus()

    txtSname.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtTel_GotFocus()

    txtTel.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtTel_LostFocus()

    txtTel.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub MakecboScode_Kana(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboScode_Kana_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        '得意先マスタ
        If Trim(strBuff1) = "" Then
            strSQL = "SELECT Bkana,Bname,Bcode FROM MT070" & _
                     " WHERE Fdiv = " & BUSINESS_DIV_EXHIBITION & " OR Fdiv = " & BUSINESS_DIV_ALL & _
                     " ORDER BY Bkana,Bname,Bcode"
        Else
            strSQL = "SELECT Bkana,Bname,Bcode FROM MT070" & _
                     " WHERE (Fdiv = " & BUSINESS_DIV_EXHIBITION & " OR Fdiv = " & BUSINESS_DIV_ALL & ")" & _
                     " AND Bkana LIKE '" & strBuff1 & "%'" & _
                     " ORDER BY Bkana,Bname,Bcode"
        End If
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Ctrl.AddItem .Fields("Bkana") & ";" & .Fields("Bname") & ";" & .Fields("Bcode")
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboScode_Kana_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboScode_Kana_Err")

End Sub

'目　的　　：合計の計算
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub Calc_Total()

    Dim intIndex1 As Integer
    Dim curTotal As Currency
    Dim varPrice As Variant

    On Error GoTo Calc_Total_Err
    
    curTotal = 0
    varPrice = 0
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        curTotal = curTotal + CCur(lsvMeisai.ListItems(intIndex1).SubItems(3))
        If Trim(lsvMeisai.ListItems(intIndex1).SubItems(7)) <> "" Then
            varPrice = varPrice + CDec(lsvMeisai.ListItems(intIndex1).SubItems(7))
        End If
    Next intIndex1
    
    imnQty_Total.Value = curTotal
    If varPrice > imnPrice_Total.MaxValue Then
        imnPrice_Total.Value = imnPrice_Total.MaxValue
    ElseIf varPrice < imnPrice_Total.MinValue Then
        imnPrice_Total.Value = imnPrice_Total.MinValue
    Else
        imnPrice_Total.Value = varPrice
    End If
    
    Exit Sub
    
Calc_Total_Err:

    Call MsgBox("合計の計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Total_Err")

End Sub

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub MakecboIcode_Kana(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboIcode_Kana_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        '商品マスタ
        If Trim(strBuff1) = "" Then
            strSQL = "SELECT Ikana,Iname,Icode FROM MT050" & _
                     " ORDER BY Ikana,Iname,Icode"
        Else
            strSQL = "SELECT Ikana,Iname,Icode FROM MT050" & _
                     " WHERE Ikana LIKE '%" & strBuff1 & "%'" & _
                     " ORDER BY Ikana,Iname,Icode"
        End If
        .Open strSQL, g_clsAdoSQL.Connection, adOpenForwardOnly, adLockReadOnly
        Do While Not .EOF
            Ctrl.AddItem .Fields("Ikana") & ";" & .Fields("Iname") & ";" & .Fields("Icode")
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboIcode_Kana_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboIcode_Kana_Err")

End Sub

'目　的　　：明細入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function DoValidationChecks_Dst() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Dst_Err
    
    If imnNo.Value > MAX_ROW Then
        strErrMsg = StrConv((MAX_ROW + 1), vbWide) & "行以上入力できません。"
        txtIname.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtIname.Text) = "" Then
        strErrMsg = "植木名を入力してください。"
        txtIname.SetFocus
        GoTo ErrorTrap:
    End If
    If imnQty.Value = 0 Then
        strErrMsg = "数量を入力してください。"
        imnQty.SetFocus
        GoTo ErrorTrap:
    End If
    If imnPrice.Value <> 0 Then
        If Trim(cboBcode.Text) = "" Then
            strErrMsg = "買主を入力してください。"
            If chkChumon.Value = 0 Then chkChumon.Value = 1
            DoEvents
            cboBcode.SetFocus
            GoTo ErrorTrap:
        End If
    End If

    DoValidationChecks_Dst = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks_Dst = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks_Dst_Err:

    DoValidationChecks_Dst = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Dst_Err")

End Function

'目　的　　：リストビューからの行番号取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function ListViewGetMaxRow() As Boolean

    On Error GoTo ListViewGetMaxRow_Err

    ListViewGetMaxRow = False

    '行番号取得
    imnNo.Value = lsvMeisai.ListItems.Count + 1

    ListViewGetMaxRow = True

    Exit Function

ListViewGetMaxRow_Err:

    Call MsgBox("リストビューからの行番号取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewGetMaxRow_Err")

End Function

'目　的　　：リストビューへのデータ挿入
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function ListViewInsItem() As Boolean
    
    Dim varRes As Variant
    Dim intPostion As Integer
    
    On Error GoTo ListViewInsItem_Err
    
    ListViewInsItem = False
    
    If lsvMeisai.ListItems.Count >= MAX_ROW Then
        Call MsgBox("これ以上明細を入力できません。", vbOKOnly + vbCritical, "")
        Exit Function
    End If
    
    varRes = InputBox("挿入する行番号を入力してください...", "", "")

    '入力値をチェック
    If Trim(varRes) = "" Then
        Call MsgBox("行番号を入力してください。", vbOKOnly + vbCritical, "")
        Exit Function
    End If
    If IsNumeric(varRes) = False Then
        Call MsgBox("行番号が不正です。", vbOKOnly + vbCritical, "")
        Exit Function
    End If

    If DoValidationChecks_Dst() = False Then Exit Function

    '編集行の取得
    intPostion = CInt(varRes)
    
    Call ListViewSetItem(intPostion, 1)

    '行番号振り直し
    Call ListViewRefresh

    ListViewInsItem = True

    Exit Function

ListViewInsItem_Err:

    Call MsgBox("リストビューへのデータ挿入エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewInsItem_Err")

End Function

'目　的　　：リストビューの行番号を振り直す
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function ListViewRefresh() As Boolean

    Dim intIndex1 As Integer

    On Error GoTo ListViewRefresh_Err

    ListViewRefresh = False

    lsvMeisai.Refresh
    For intIndex1 = 1 To lsvMeisai.ListItems.Count Step 1
        lsvMeisai.ListItems(intIndex1).Text = intIndex1
    Next intIndex1

    ListViewRefresh = True

    Exit Function

ListViewRefresh_Err:

    Call MsgBox("リストビューの行番号を振り直しエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewRefresh_Err")

End Function

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Sub MakecboScode(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboScode_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        '得意先マスタ
        strSQL = "SELECT Bkana,Bname,Bcode FROM MT070" & _
                 " WHERE Fdiv = " & BUSINESS_DIV_EXHIBITION & " OR Fdiv = " & BUSINESS_DIV_ALL & _
                 " ORDER BY Bcode"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboScode_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboScode_Err")

End Sub

'目　的　　：売立金額の計算
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Private Function Calc_Price() As Boolean

    Dim varPrice As Variant

    On Error GoTo Calc_Price_Err
    
    Calc_Price = False
    
    varPrice = CDec(imnQty.Value) * CDec(imnPrice1.Value)
    If varPrice > imnPrice.MaxValue Then
        Call MsgBox("売立金額が大きすぎます。", vbOKOnly + vbCritical, "")
        DoEvents
    ElseIf varPrice < imnPrice.MinValue Then
        Call MsgBox("売立金額が小さすぎます。", vbOKOnly + vbCritical, "")
        DoEvents
    Else
        imnPrice.Value = CDec(imnQty.Value) * CDec(imnPrice1.Value)
        Calc_Price = True
    End If
    
    Exit Function
    
Calc_Price_Err:

    Call MsgBox("売立金額の計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Price_Err")

End Function

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
'更新履歴　：
'
Private Sub MakecboBcode(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboBcode_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        strSQL = "SELECT * FROM vw_MT071" & _
                 " WHERE (Fdiv = " & BUSINESS_DIV_BUYER & " OR Fdiv = " & BUSINESS_DIV_ALL & ")"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            If Not IsNull(.Fields("Sdate")) And Not IsNull(.Fields("Fdate")) Then
                If .Fields("Sdate") <= Trim(lblOdate.Caption) And Trim(lblOdate.Caption) <= .Fields("Fdate") Then
                    Ctrl.AddItem .Fields("Bnum") & ";" & .Fields("Bname")
                Else
                    Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
                End If
            Else
                Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
            End If
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboBcode_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboBcode_Err")

End Sub

Private Function MakeWork() As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim intPage As Integer
    Dim intLineCount As Integer
    Dim intIndex1 As Integer
    
    Const PAGE_MAX_ROW = 10
    
    On Error GoTo MakeWork_Err
    
    MakeWork = False
    
    Screen.MousePointer = vbHourglass
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF012"
    g_clsAdoAccess.Connection.Execute strSQL
    
    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF012"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    'データオープン
    strSQL = "SELECT * FROM DT011" & _
             " WHERE Odate = '" & frmYpmf010.lblOdate.Caption & "'" & _
             " AND Pnum = " & txtPnum.Text & _
             " ORDER BY Odate,Pnum,Line"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    
    intPage = 1
    intLineCount = 0
        
    Do While Not adoRecordset1.EOF
        intLineCount = intLineCount + 1
        If intLineCount > PAGE_MAX_ROW Then
            intPage = intPage + 1
            intLineCount = 1
        End If
        
        wkRecordset.AddNew
        wkRecordset.Fields("Odate") = adoRecordset1.Fields("Odate")
        wkRecordset.Fields("Pnum") = adoRecordset1.Fields("Pnum")
        wkRecordset.Fields("Pagenum") = intPage
        wkRecordset.Fields("No") = intLineCount
        wkRecordset.Fields("Line") = adoRecordset1.Fields("Line")
        wkRecordset.Fields("Icode") = adoRecordset1.Fields("Icode")
        wkRecordset.Fields("Iname") = adoRecordset1.Fields("Iname")
        wkRecordset.Fields("Qty") = adoRecordset1.Fields("Qty")
        wkRecordset.Update
        
        adoRecordset1.MoveNext
    Loop
        
    If intLineCount < PAGE_MAX_ROW Then
        '空行作成
        For intIndex1 = intLineCount + 1 To PAGE_MAX_ROW
            wkRecordset.AddNew
            wkRecordset.Fields("Odate") = frmYpmf010.lblOdate.Caption
            wkRecordset.Fields("Pnum") = Null
            wkRecordset.Fields("Pagenum") = intPage
            wkRecordset.Fields("No") = intIndex1
            wkRecordset.Update
        Next intIndex1
    End If
        
    adoRecordset1.Close
    wkRecordset.Requery
    wkRecordset.Close
    
    Screen.MousePointer = vbDefault
    
    MakeWork = True
    
    Exit Function
    
MakeWork_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    Screen.MousePointer = vbDefault
    Call MsgBox("ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork_Err")
    
End Function


