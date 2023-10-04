VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchDT011 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "受付表から検索"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14670
   Icon            =   "frmSearchDT011.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdRemove 
      Caption         =   "戻す(F 9)<<"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1860
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   7560
   End
   Begin VB.Frame Frame2 
      Caption         =   "行選択"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4380
      TabIndex        =   19
      Top             =   60
      Width           =   2295
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "受付行番号"
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
      Begin imText6Ctl.imText imtPnumLine 
         Height          =   465
         Left            =   1620
         TabIndex        =   3
         Top             =   300
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   820
         Caption         =   "frmSearchDT011.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSearchDT011.frx":007A
         Key             =   "frmSearchDT011.frx":0098
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
         AutoConvert     =   0
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "検索条件"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Width           =   4155
      Begin VB.CommandButton cmdSearch 
         Caption         =   "検索開始"
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
         Left            =   2700
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   1335
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   18
         Top             =   300
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
      Begin imText6Ctl.imText imtPnum 
         Height          =   465
         Left            =   1620
         TabIndex        =   1
         Top             =   300
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   820
         Caption         =   "frmSearchDT011.frx":00CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSearchDT011.frx":013A
         Key             =   "frmSearchDT011.frx":0158
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
         AutoConvert     =   0
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
   End
   Begin VB.TextBox txtKey 
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
      Index           =   0
      Left            =   2340
      TabIndex        =   16
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtKey 
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
      Index           =   1
      Left            =   3060
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtKey 
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
      Index           =   2
      Left            =   3780
      TabIndex        =   14
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtKey 
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
      Index           =   3
      Left            =   4500
      TabIndex        =   13
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "キャンセル(&C)"
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
      Left            =   7560
      TabIndex        =   10
      Top             =   7560
      Width           =   1875
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "OK(&O)"
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
      Left            =   5580
      TabIndex        =   9
      Top             =   7560
      Width           =   1875
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   120
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=YAWASESRC"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "YAWASESRC"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM DT011 ORDER BY Odate,Pnum,Linr"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   14700
      TabIndex        =   0
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmSearchDT011.frx":018C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011.frx":01FA
      Key             =   "frmSearchDT011.frx":0218
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
      Left            =   14880
      TabIndex        =   11
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmSearchDT011.frx":025C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011.frx":02CA
      Key             =   "frmSearchDT011.frx":02E8
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
   Begin MSComctlLib.ListView lsvMeisai 
      Height          =   6435
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   11351
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "行"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "商品コード"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "植木名称"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "数 量"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "合算(F11)>>"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddnew 
      BackColor       =   &H00FFFFC0&
      Caption         =   "受付新規登録(F12)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7260
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   2115
   End
   Begin MSComctlLib.ListView lsvTotal 
      Height          =   6735
      Left            =   11100
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   11880
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "行"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "商品コード"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "植木名称"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "数 量"
         Object.Width           =   2822
      EndProperty
   End
   Begin imText6Ctl.imText imtUnloadDummy 
      Height          =   165
      Left            =   14700
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   540
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   291
      Caption         =   "frmSearchDT011.frx":032C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011.frx":039A
      Key             =   "frmSearchDT011.frx":03B8
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
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   2
      LengthAsByte    =   0
      Text            =   ""
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
   Begin VB.Label lblDummy 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   6435
      Left            =   120
      TabIndex        =   12
      Top             =   1020
      Width           =   9315
   End
   Begin VB.Menu mnuPop 
      Caption         =   "ポップアップメニュー"
      Visible         =   0   'False
      Begin VB.Menu mnuAsc 
         Caption         =   "昇　順"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesc 
         Caption         =   "降　順"
      End
   End
End
Attribute VB_Name = "frmSearchDT011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_intAddNew_PnumLine As Integer  '新規登録時の行番号
Public g_curPrice As Currency           '明細合算時の売立金額
Public g_strBcode As String             '明細合算時の買主コード
Public g_strBname As String             '明細合算時の買主名称

Private Const KEY_COUNT = 4             'キーの数
Private Const DT021_MAX_ROW = 20        '競売結果の明細数
Private Const DT011_MAX_ROW = 20        '受付データの明細数

Private Sub cmdAddnew_Click()

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo cmdAddnew_Click_Err

    g_intAddNew_PnumLine = 0

    '入力チェック
    If Trim(imtPnum.Text) = "" Then
        imtPnum.SetFocus
        DoEvents
        Call MsgBox("受付番号を入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If

    '受付データ
    strSQL = "SELECT * FROM DT011" & _
             " WHERE Odate = '" & frmYpmf020.lblOdate & "'" & _
             " AND Pnum = " & imtPnum.Text & _
             " ORDER BY Odate,Pnum,Line DESC"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = True Then
        imtPnum.SetFocus
        DoEvents
        Call MsgBox("受付データがありません。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    If adoRecordset1.Fields("Line") >= DT011_MAX_ROW Then
        imtPnum.SetFocus
        DoEvents
        Call MsgBox("これ以上受付データを追加することはできません。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    
    g_intAddNew_PnumLine = CInt(adoRecordset1.Fields("Line")) + 1
    adoRecordset1.Close
    frmSearchDT011AddNew.Show vbModal

    Exit Sub

cmdAddnew_Click_Err:

    Call MsgBox("新規登録クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdAddnew_Click_Err")

End Sub

Private Sub cmdAddnew_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Form_KeyDown(KeyCode, Shift)
    
End Sub

Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If Trim(imtPnum.Text) = "" Then
        imtPnum.SetFocus
        DoEvents
        Call MsgBox("受付番号を入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
            
    If lsvTotal.ListItems.Count <= 0 Then
        '通常選択
        If SetData() = False Then Exit Sub
    Else
        '合算
        g_curPrice = 0
        frmSearchDT011Total.Show vbModal
        If g_curPrice <> 0 Then
            If SetDataMulti() = False Then Exit Sub
        End If
    End If
    
    Exit Sub
    
cmdExecute_Click_Err:

    Call MsgBox("ＯＫクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")
    
End Sub

Private Sub cmdExecute_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Form_KeyDown(KeyCode, Shift)
    
End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdExit_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    Call Form_KeyDown(KeyCode, Shift)

End Sub

Private Sub cmdRemove_Click()

    Dim itmX As ListItem
    Dim intIndex1 As Integer

    On Error GoTo cmdTotal_Click_Err

    '明細データへ追加
    For intIndex1 = 1 To lsvTotal.ListItems.Count
        If lsvTotal.ListItems(intIndex1).Selected = True Then
            Set itmX = lsvMeisai.ListItems.Add(, , lsvTotal.ListItems(intIndex1).Text, 0)
            itmX.SubItems(1) = lsvTotal.ListItems(intIndex1).SubItems(1)
            itmX.SubItems(2) = lsvTotal.ListItems(intIndex1).SubItems(2)
            itmX.SubItems(3) = lsvTotal.ListItems(intIndex1).SubItems(3)
        End If
    Next intIndex1

    '行番号でソート
    lsvMeisai.SortKey = 0
    lsvMeisai.SortOrder = lvwAscending
    lsvMeisai.Sorted = True
    lsvMeisai.Refresh

    '明細データから合算データから削除
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        'リストビューのデータ検索（行番号が一致するデータがあったら削除）
        Set itmX = lsvTotal.FindItem(lsvMeisai.ListItems(intIndex1).Text, , , 0)
        If Not (itmX Is Nothing) Then
            'データ削除
            lsvTotal.ListItems.Remove itmX.Index
        End If
    Next intIndex1

    Exit Sub

cmdTotal_Click_Err:

    Call MsgBox("合算クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdTotal_Click_Err")

End Sub

Private Sub cmdRemove_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Form_KeyDown(KeyCode, Shift)
    
End Sub

Private Sub cmdSearch_Click()

    On Error GoTo cmdSearch_Click_Err

    If Trim(imtPnum.Text) = "" Then
        imtPnum.SetFocus
        DoEvents
        Call MsgBox("受付番号を入力してください。", vbOKOnly + vbCritical, "")
        Exit Sub
    End If
    
    If SearchData() = False Then Exit Sub

    Exit Sub

cmdSearch_Click_Err:

    Call MsgBox("検索開始クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdSearch_Click_Err")

End Sub

Private Sub cmdSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Form_KeyDown(KeyCode, Shift)

End Sub

Private Sub cmdTotal_Click()
    
    Dim intLineRemainder As Integer     '残り行数
    Dim itmX As ListItem
    Dim intIndex1 As Integer

    On Error GoTo cmdTotal_Click_Err

    '入力チェック
    intLineRemainder = DT021_MAX_ROW - CInt(frmYpmf020.lsvMeisai.ListItems.Count)
    If intLineRemainder < 2 Then
        Call MsgBox("行数が足りないため合算できません。", vbOKOnly + vbInformation, "")
        Exit Sub
    End If

    '合算データへ追加
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        If lsvMeisai.ListItems(intIndex1).Selected = True Then
            Set itmX = lsvTotal.ListItems.Add(, , lsvMeisai.ListItems(intIndex1).Text, 0)
            itmX.SubItems(1) = lsvMeisai.ListItems(intIndex1).SubItems(1)
            itmX.SubItems(2) = lsvMeisai.ListItems(intIndex1).SubItems(2)
            itmX.SubItems(3) = lsvMeisai.ListItems(intIndex1).SubItems(3)
            intLineRemainder = intLineRemainder - 1
        End If
        If intLineRemainder <= 0 Then Exit For
    Next intIndex1

    '行番号でソート
    lsvTotal.SortKey = 0
    lsvTotal.SortOrder = lvwAscending
    lsvTotal.Sorted = True
    lsvTotal.Refresh

    '合算データから明細データを削除
    For intIndex1 = 1 To lsvTotal.ListItems.Count
        'リストビューのデータ検索（行番号が一致するデータがあったら削除）
        Set itmX = lsvMeisai.FindItem(lsvTotal.ListItems(intIndex1).Text, , , 0)
        If Not (itmX Is Nothing) Then
            'データ削除
            lsvMeisai.ListItems.Remove itmX.Index
        End If
    Next intIndex1

    Exit Sub

cmdTotal_Click_Err:

    Call MsgBox("合算クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdTotal_Click_Err")

End Sub

Private Sub cmdTotal_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Form_KeyDown(KeyCode, Shift)
    
End Sub

Private Sub imtPnum_GotFocus()
    
    imtPnum.BackColor = FOCUS_STOP_COLOR
    imtPnum.Tag = imtPnum.Text
    
End Sub

Private Sub imtPnum_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Form_KeyDown(KeyCode, Shift)

End Sub

Private Sub imtPnum_LostFocus()

    imtPnum.BackColor = FOCUS_NO_COLOR
    imtPnum.Tag = ""
    
End Sub

Private Sub imtPnum_Validate(Cancel As Boolean)

    On Error GoTo imtPnum_Validate_Err

    If imtPnum.Tag = imtPnum.Text Then Exit Sub
    If Trim(imtPnum.Text) = "" Then
        lsvMeisai.ListItems.Clear
        lsvMeisai.Visible = False
        lblDummy.Visible = True
        Exit Sub
    End If

    If SearchData() = False Then Cancel = True
    
    imtPnum.Tag = ""
    
    Exit Sub

imtPnum_Validate_Err:
    
    Call MsgBox("受付番号フォーカス移動時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtPnum_Validate_Err")

End Sub

Private Sub imtPnumLine_GotFocus()
    
    imtPnumLine.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imtPnumLine_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Form_KeyDown(KeyCode, Shift)

End Sub

Private Sub imtPnumLine_LostFocus()
    
    On Error GoTo imtPnumLine_LostFocus_Err

    imtPnumLine.BackColor = FOCUS_NO_COLOR

    If imtPnumLine.Tag = "True" Then
        Call SetData
    End If

    Exit Sub

imtPnumLine_LostFocus_Err:
    
    Call MsgBox("受付行番号フォーカス喪失時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtPnumLine_LostFocus_Err")

End Sub

Private Sub imtPnumLine_Validate(Cancel As Boolean)
    
    Dim intIndex1 As Integer

    On Error GoTo imtPnumLine_Validate_Err

    If Trim(imtPnumLine.Text) = "" Then Exit Sub
    
    imtPnumLine.Tag = ""
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        lsvMeisai.ListItems(intIndex1).Selected = False
        If CInt(imtPnumLine.Text) = CInt(lsvMeisai.ListItems(intIndex1).Text) Then
            lsvMeisai.ListItems(intIndex1).Selected = True
            
            txtKey(0).Text = imtPnumLine.Text
            txtKey(1).Text = lsvMeisai.ListItems(intIndex1).SubItems(1)
            txtKey(2).Text = lsvMeisai.ListItems(intIndex1).SubItems(2)
            txtKey(3).Text = lsvMeisai.ListItems(intIndex1).SubItems(3)
            imtPnumLine.Tag = "True"
            Exit Sub
        End If
    Next intIndex1

    Cancel = True
    Call MsgBox("受付行番号が存在しません。", vbOKOnly + vbCritical, "")

    Exit Sub

imtPnumLine_Validate_Err:
    
    Call MsgBox("受付行番号フォーカス移動時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "imtPnumLine_Validate_Err")

End Sub

Private Sub imtUnloadDummy_GotFocus()

    Unload Me

End Sub

Private Sub lsvMeisai_Click()

    On Error GoTo lsvMeisai_Click_Err

    '行が選択されているか？
    If lsvMeisai.SelectedItem Is Nothing Then Exit Sub

    txtKey(0).Text = lsvMeisai.SelectedItem.Text
    txtKey(1).Text = lsvMeisai.SelectedItem.SubItems(1)
    txtKey(2).Text = lsvMeisai.SelectedItem.SubItems(2)
    txtKey(3).Text = lsvMeisai.SelectedItem.SubItems(3)
    
    Exit Sub
    
lsvMeisai_Click_Err:
    
    Call MsgBox("明細クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "lsvMeisai_Click_Err")
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Form_KeyDown_Err
    
    If KeyCode = vbKeyEscape Then
        imtUnloadDummy.SetFocus
        DoEvents
        Exit Sub
    End If
    
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
            cmdRemove.SetFocus
            DoEvents
            Call cmdRemove_Click
        Case vbKeyF10
        Case vbKeyF11
            cmdTotal.SetFocus
            DoEvents
            Call cmdTotal_Click
        Case vbKeyF12
'            cmdAddnew.SetFocus
'            DoEvents
'            Call cmdAddnew_Click
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

Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Call FieldsClear

    imtPnum.Text = frmYpmf020.imtPnum.Text
    Timer1.Enabled = True

    Exit Sub

Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    imtPnum.SetFocus

End Sub

Private Sub lsvMeisai_DblClick()

    On Error GoTo lsvMeisai_DblClick_Err

    Call lsvMeisai_Click
    Call cmdExecute_Click
    
    Exit Sub
    
lsvMeisai_DblClick_Err:
    
    Call MsgBox("明細ダブルクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "lsvMeisai_DblClick_Err")
  
End Sub

Private Function SearchData() As Boolean

    Dim strSQL As String
    Dim strWhere As String
    Dim itmX As ListItem
    Dim blnFlg As Boolean
    Dim intIndex1 As Integer
    Dim curQty As Currency
    Dim strDeleteData As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo SearchData_Err

    SearchData = False

    Screen.MousePointer = vbHourglass

'    strSQL = "SELECT * FROM DT011 "
'    strWhere = " Odate = '" & frmYpmf020.lblOdate.Caption & "'"
'    strWhere = strWhere & " AND Pnum = " & imtPnum.Text
'    strWhere = strWhere & " AND (Idiv IS NULL OR Idiv <> 1)"
'    strWhere = strWhere & " AND (Price IS NULL OR Price = 0)"
    
    strSQL = "SELECT * FROM DT011 "
    strWhere = " (Odate = '" & frmYpmf020.lblOdate.Caption & "'"
    strWhere = strWhere & " AND Pnum = " & imtPnum.Text
    strWhere = strWhere & " AND (Idiv IS NULL OR Idiv <> 1)"
    strWhere = strWhere & " AND (Price IS NULL OR Price = 0) )"
    
    '変更処理時のみ(2004/01/31追加)
    If frmYpmf020.optSyori(1).Value = True Then
        If UBound(g_usrMeisaiDel) >= 1 Then
            '削除ワーク
            strDeleteData = ""
            For intIndex1 = 1 To UBound(g_usrMeisaiDel)
                If Trim(imtPnum.Text) = Trim(g_usrMeisaiDel(intIndex1).Pnum) Then
                    If strDeleteData = "" Then
                        strWhere = strWhere & " OR (Odate = '" & frmYpmf020.lblOdate.Caption & "'"
                        strWhere = strWhere & " AND ("
                    Else
                        strDeleteData = strDeleteData & " OR "
                    End If
                    strDeleteData = strDeleteData & " (Pnum = " & g_usrMeisaiDel(intIndex1).Pnum
                    strDeleteData = strDeleteData & " AND Line = " & g_usrMeisaiDel(intIndex1).PnumLine & ")"
                End If
            Next intIndex1
            If Trim(strDeleteData) <> "" Then
                strWhere = strWhere & strDeleteData & "))"
            End If
        End If
    End If
    

    If Trim(strWhere) <> "" Then
        Adodc1.RecordSource = strSQL & " WHERE " & strWhere
    Else
        Adodc1.RecordSource = strSQL
    End If
    
    '検索実行
    Adodc1.Refresh
    lsvMeisai.ListItems.Clear

    If Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
        'データなし
        lsvMeisai.Visible = False
        lblDummy.Visible = True
        imtPnum.SetFocus
        DoEvents
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "情報")
    Else
        lsvMeisai.Visible = True
        lblDummy.Visible = False
        
        Do While Not Adodc1.Recordset.EOF
            '既に明細に入力されているデータは除く
            blnFlg = True
            curQty = 0
            For intIndex1 = 1 To frmYpmf020.lsvMeisai.ListItems.Count
                If Trim(imtPnum.Text) = Trim(frmYpmf020.lsvMeisai.ListItems(intIndex1).SubItems(1)) Then
                    If CInt(Adodc1.Recordset.Fields("Line")) = CInt(frmYpmf020.lsvMeisai.ListItems(intIndex1).SubItems(2)) Then
                        blnFlg = False
                        '数量を集計
                        curQty = curQty + CCur(frmYpmf020.lsvMeisai.ListItems(intIndex1).SubItems(5))
                    End If
                End If
            Next
                    
                    
            '2005/09/16 入力中ワークを探す。入力中の場合は表示しない
            If blnFlg = True Then
                strSQL = "SELECT * FROM YPMF020" & _
                         " WHERE Odate = '" & g_strOdate & "'" & _
                         " AND Pnum = " & imtPnum.Text & _
                         " AND Line = " & Adodc1.Recordset.Fields("Line")
                adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
                If adoRecordset1.EOF = False Then
                    blnFlg = False
                    curQty = Adodc1.Recordset.Fields("Qty")
                End If
                adoRecordset1.Close
            End If
            
            
            '既に明細に入力されていても数量が少ない場合も入力可
            If blnFlg = True Or (blnFlg = False And curQty < CCur(Adodc1.Recordset.Fields("Qty"))) Then
                'データを追加
                Set itmX = lsvMeisai.ListItems.Add(, , Format(Adodc1.Recordset.Fields("Line"), "00"), 0)
                itmX.SubItems(1) = IIf(IsNull(Adodc1.Recordset.Fields("Icode")), "", Adodc1.Recordset.Fields("Icode"))
                itmX.SubItems(2) = IIf(IsNull(Adodc1.Recordset.Fields("Iname")), "", Adodc1.Recordset.Fields("Iname"))
                itmX.SubItems(3) = Format(CCur(Adodc1.Recordset.Fields("Qty")) - curQty, "#,##0")
            End If
            
            Adodc1.Recordset.MoveNext
        Loop
            
        If lsvMeisai.ListItems.Count <= 0 Then
            lsvMeisai.Visible = False
            lblDummy.Visible = True
            imtPnum.SetFocus
            DoEvents
            Call MsgBox("データがありません！！", vbOKOnly + vbInformation, "情報")
        Else
            SearchData = True
        End If
    End If

    Screen.MousePointer = vbDefault

    Exit Function

SearchData_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("検索開始クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "SearchData_Err")

End Function

Private Function SetData() As Boolean

    Dim intIndex1 As Integer
    Dim intSelect As Integer
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo SetData_Err

    SetData = False

    '一番最初に選択されているデータを探す
    intSelect = 0
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        If lsvMeisai.ListItems(intIndex1).Selected = True Then
            intSelect = intIndex1
            Exit For
        End If
    Next intIndex1

    If intSelect > 0 Then
        
        '2005/09/16 入力中ワークを探す。入力中の場合は入力させない
        strSQL = "SELECT * FROM YPMF020" & _
                 " WHERE Odate = '" & g_strOdate & "'" & _
                 " AND Pnum = " & imtPnum.Text & _
                 " AND Line = " & lsvMeisai.ListItems(intSelect).Text
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoRecordset1.EOF = False Then
            adoRecordset1.Close
            Call MsgBox("既に他の人が入力中です。別の行を選択して下さい。", vbOKOnly + vbCritical, "SetData_Err")
            Exit Function
        End If
        adoRecordset1.Close
    
        frmYpmf020.imtPnum.Text = imtPnum.Text
        frmYpmf020.imtPnumLine.Text = lsvMeisai.ListItems(intSelect).Text
        frmYpmf020.cboIcode.Text = lsvMeisai.ListItems(intSelect).SubItems(1)
        frmYpmf020.txtIname.Text = lsvMeisai.ListItems(intSelect).SubItems(2)
        frmYpmf020.imnQty.Value = lsvMeisai.ListItems(intSelect).SubItems(3)
        
        Unload Me
    End If
    
    SetData = True
    
    Exit Function
    
SetData_Err:

    Call MsgBox("データセット時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "SetData_Err")
    
End Function

Private Sub FieldsClear()

    Dim intIndex1 As Integer

    On Error GoTo FieldsClear_Err
    
    imtPnum.Text = ""
    imtPnumLine.Text = ""
    For intIndex1 = 0 To KEY_COUNT - 1
        txtKey(intIndex1).Text = ""
    Next intIndex1
    lsvMeisai.ListItems.Clear
    lsvMeisai.Visible = False
    lblDummy.Visible = True
    lsvTotal.ListItems.Clear
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

Private Sub lsvMeisai_GotFocus()

    If lsvMeisai.ListItems.Count > 0 Then
        lsvMeisai.ListItems(1).Selected = True
        Call lsvMeisai_Click
    End If

End Sub

Private Sub lsvMeisai_ItemClick(ByVal Item As MSComctlLib.ListItem)

'    Call lsvMeisai_Click

End Sub

Private Sub lsvMeisai_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn And Shift <> vbShiftMask Then
        KeyCode = 0
        Call lsvMeisai_DblClick
        Exit Sub
    End If

    Call Form_KeyDown(KeyCode, Shift)

End Sub

Private Sub Timer1_Timer()

    On Error GoTo Timer1_Timer_Err

    Timer1.Enabled = False

    If Trim(imtPnum.Text) <> "" Then
        If SearchData() = False Then Exit Sub
        imtPnumLine.SetFocus
    Else
        imtPnum.SetFocus
    End If

    Exit Sub

Timer1_Timer_Err:

    Call MsgBox("タイマーエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Timer1_Timer_Err")

End Sub

Private Function SetDataMulti() As Boolean
    
    Dim intIndex1 As Integer
    Dim itmX As ListItem
    Dim intPostion As Integer
    Dim blnFlg As Boolean           '合算明細
    Dim intPnumLine As Integer      '合算行番号
    
    On Error GoTo SetDataMulti_Err

    SetDataMulti = False

    blnFlg = True
    intPnumLine = 0
    intPostion = frmYpmf020.lsvMeisai.ListItems.Count + 1
    
    For intIndex1 = 1 To lsvTotal.ListItems.Count
        '合算行番号取得
        If intPnumLine = 0 Then intPnumLine = CInt(lsvTotal.ListItems(intIndex1).Text)
        
        'メイン画面の明細データへ追加
        Set itmX = frmYpmf020.lsvMeisai.ListItems.Add(, , intPostion, 0)
        itmX.SubItems(1) = Trim(imtPnum.Text)
        itmX.SubItems(2) = lsvTotal.ListItems(intIndex1).Text
        itmX.SubItems(3) = lsvTotal.ListItems(intIndex1).SubItems(1)
        itmX.SubItems(4) = lsvTotal.ListItems(intIndex1).SubItems(2)
        itmX.SubItems(5) = lsvTotal.ListItems(intIndex1).SubItems(3)
        If blnFlg = True Then
            itmX.SubItems(6) = Format(g_curPrice, "#,##0")
            blnFlg = False
        Else
            itmX.SubItems(6) = "0"
        End If
        itmX.SubItems(7) = g_strBcode
        itmX.SubItems(8) = g_strBname
'            itmX.SubItems(9) = chkWdiv.Value
'            itmX.SubItems(10) = chkSdiv.Value
'            itmX.SubItems(11) = chkBdiv.Value
'            itmX.SubItems(12) = imnBnum.Value
'            itmX.SubItems(13) = imnSnum.Value
'            itmX.SubItems(14) = Trim(lblItime.Caption)
'            itmX.SubItems(15) = Trim(lblDetailPcode.Caption)
'            itmX.SubItems(16) = Trim(lblDetailPname.Caption)
        itmX.SubItems(17) = intPnumLine
        itmX.SubItems(18) = AUCTION_ON
        If Trim(itmX.SubItems(17)) <> "" And itmX.SubItems(17) <> "0" Then
            itmX.SubItems(19) = "合"
        End If
        If Trim(itmX.SubItems(18)) <> "" And itmX.SubItems(18) <> "0" Then
            itmX.SubItems(19) = "ﾔﾒ"
        End If
        
        intPostion = intPostion + 1
    Next intIndex1
    
    'メイン画面の明細クリア
    frmYpmf020.imnNo.Value = intPostion
'    frmYpmf020.imtPnum.Text = ""
    frmYpmf020.imtPnumLine.Text = ""
    frmYpmf020.cboIcode.Text = ""
    frmYpmf020.txtIname.Text = ""
    frmYpmf020.imnQty.Value = ""
    frmYpmf020.imnPrice.Value = 0
    
    Call frmYpmf020.Calc_Total      '合計計算
    
    frmYpmf020.m_blnTotalFlg = True
    Unload Me
    
    SetDataMulti = True
    
    Exit Function
    
SetDataMulti_Err:

    Call MsgBox("データセット時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "SetDataMulti_Err")
    
End Function

