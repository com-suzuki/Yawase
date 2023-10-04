VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "競売結果検索"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdPrint 
      Caption         =   "印　刷"
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
      Left            =   60
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8340
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtKey1 
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
      Left            =   6000
      TabIndex        =   8
      Top             =   8340
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "キャンセル"
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
      Left            =   8640
      TabIndex        =   5
      Top             =   8340
      Width           =   1875
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "OK"
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
      Left            =   6660
      TabIndex        =   4
      Top             =   8340
      Width           =   1875
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   1620
      Top             =   8340
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
      RecordSource    =   "SELECT * FROM DT020 ORDER BY Odate,Ocode"
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
      Left            =   11400
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmSearch.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearch.frx":007A
      Key             =   "frmSearch.frx":0098
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
      Left            =   11640
      TabIndex        =   7
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmSearch.frx":00DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearch.frx":014A
      Key             =   "frmSearch.frx":0168
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   435
      Left            =   3780
      Top             =   8340
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
      RecordSource    =   "SELECT * FROM DT021 ORDER BY Ocode,Pnum"
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
   Begin VB.Frame fraSearch1 
      Height          =   7455
      Left            =   120
      TabIndex        =   10
      Top             =   540
      Width           =   10275
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
         Height          =   795
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   10035
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
            Left            =   8040
            TabIndex        =   3
            Top             =   240
            Width           =   1875
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Width           =   1455
            _Version        =   262145
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "競売番号"
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
            Index           =   1
            Left            =   4140
            TabIndex        =   14
            Top             =   300
            Width           =   1455
            _Version        =   262145
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "手板番号"
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
         Begin imText6Ctl.imText txtHnum 
            Height          =   360
            Left            =   5640
            TabIndex        =   2
            Top             =   300
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   635
            Caption         =   "frmSearch.frx":01AC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSearch.frx":021A
            Key             =   "frmSearch.frx":0238
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
         Begin imText6Ctl.imText txtOcode 
            Height          =   360
            Left            =   1620
            TabIndex        =   1
            Top             =   300
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   635
            Caption         =   "frmSearch.frx":026C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSearch.frx":02DA
            Key             =   "frmSearch.frx":02F8
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
            MaxLength       =   12
            LengthAsByte    =   0
            Text            =   "99999999999"
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
            Caption         =   "～"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   15
            Top             =   360
            Width           =   375
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmSearch.frx":033C
         Height          =   6315
         Left            =   120
         TabIndex        =   11
         Top             =   1020
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   11139
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   18
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Ocode"
            Caption         =   "競売番号"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1041
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Hnum"
            Caption         =   "手板番号"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1041
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            SizeMode        =   1
            Locked          =   -1  'True
            BeginProperty Column00 
               Alignment       =   1
               ColumnWidth     =   2220.094
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2684.977
            EndProperty
         EndProperty
      End
      Begin VB.Label lblDummy 
         Appearance      =   0  'ﾌﾗｯﾄ
         BorderStyle     =   1  '実線
         ForeColor       =   &H80000008&
         Height          =   6315
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   10035
      End
   End
   Begin VB.Frame fraSearch2 
      Height          =   7455
      Left            =   120
      TabIndex        =   17
      Top             =   540
      Width           =   10275
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmSearch.frx":0351
         Height          =   6315
         Left            =   120
         TabIndex        =   24
         Top             =   1020
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   11139
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   18
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Pnum"
            Caption         =   "受付番号"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1041
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "PnumLine"
            Caption         =   "行"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1041
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Iname"
            Caption         =   "植木名称"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1041
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Price"
            Caption         =   "売立金額"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1041
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Bcode"
            Caption         =   "買主"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1041
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Ocode"
            Caption         =   "競売番号"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1041
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            SizeMode        =   1
            Locked          =   -1  'True
            BeginProperty Column00 
               Alignment       =   1
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4155.024
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1544.882
            EndProperty
         EndProperty
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
         Height          =   795
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   10035
         Begin VB.CommandButton cmdSearch2 
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
            Left            =   8040
            TabIndex        =   23
            Top             =   240
            Width           =   1875
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
         Begin imText6Ctl.imText txtPnum 
            Height          =   360
            Left            =   1320
            TabIndex        =   20
            Top             =   300
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
            _ExtentY        =   635
            Caption         =   "frmSearch.frx":0366
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSearch.frx":03D4
            Key             =   "frmSearch.frx":03F2
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
            Index           =   2
            Left            =   2220
            TabIndex        =   26
            Top             =   300
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "買主番号"
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
         Begin imText6Ctl.imText txtBcode 
            Height          =   360
            Left            =   3300
            TabIndex        =   21
            Top             =   300
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
            _ExtentY        =   635
            Caption         =   "frmSearch.frx":0436
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSearch.frx":04A4
            Key             =   "frmSearch.frx":04C2
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
            Index           =   4
            Left            =   4200
            TabIndex        =   27
            Top             =   300
            Width           =   855
            _Version        =   262145
            _ExtentX        =   1508
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
            LabelLeft       =   6
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
         Begin imText6Ctl.imText txtIname 
            Height          =   345
            Left            =   5100
            TabIndex        =   22
            Top             =   300
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   609
            Caption         =   "frmSearch.frx":0506
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSearch.frx":0574
            Key             =   "frmSearch.frx":0592
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
            Text            =   "WWWWWWWWWW"
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   4
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
      End
      Begin VB.Label lblDummy2 
         Appearance      =   0  'ﾌﾗｯﾄ
         BorderStyle     =   1  '実線
         ForeColor       =   &H80000008&
         Height          =   6315
         Left            =   120
         TabIndex        =   25
         Top             =   1020
         Width           =   10035
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8175
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14420
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "手板番号検索"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "受付番号と買主検索"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_intHeadClickCol As Integer            'ヘッダークリック列

Public objArPrint As New clsArPrint

Private Sub cmdExecute_Click()

    If Trim(txtKey1.Text) = "" Then
        Call MsgBox("データを選択してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    
    frmYpmf020.txtOcode.Text = txtKey1.Text
    Call frmYpmf020.FieldsSet(True)
    Unload Me
    
End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdPrint_Click()

    Dim objRpt As New rptYpmf020S
    
    On Error GoTo cmdPrint_Click_Err
    
    If Adodc2.Recordset.BOF Or Adodc2.Recordset.EOF Then
        DoEvents
        Call MsgBox("データがありません！！", vbOKOnly + vbInformation, "情報")
        DoEvents
        Exit Sub
    End If
    If DataGrid2.Visible = False Then
        DoEvents
        Call MsgBox("データがありません！！", vbOKOnly + vbInformation, "情報")
        DoEvents
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "競売結果確認表"
        .objReport = objRpt
        .Connection = g_clsAdoSQL.Connection
        .SQL = Adodc2.RecordSource
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
    
cmdPrint_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("印刷クリックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdPrint_Click_Err")

End Sub

Private Sub cmdSearch_Click()

    Dim strSQL As String
    Dim strWhere As String

    On Error GoTo cmdSearch_Click_Err

    Screen.MousePointer = vbHourglass

    cmdSearch.SetFocus

    strSQL = "SELECT * FROM DT020 "
    strWhere = " Odate = '" & frmYpmf020.lblOdate.Caption & "'"

    If Trim(txtOcode.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " Ocode >= '" & txtOcode.Text & "'"
    End If
    If Trim(txtHnum.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " Hnum LIKE '%" & txtHnum.Text & "%'"
    End If

    '検索実行
    If Trim(strWhere) <> "" Then
        Adodc1.RecordSource = strSQL & " WHERE " & strWhere
    Else
        Adodc1.RecordSource = strSQL
    End If
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1

    If Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
        DataGrid1.Visible = False
        lblDummy.Visible = True
        Call MsgBox("データがありません！！", vbOKOnly + vbInformation, "情報")
    Else
        DataGrid1.Visible = True
        lblDummy.Visible = False
    End If

    Screen.MousePointer = vbDefault

    Exit Sub

cmdSearch_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("検索開始クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdSearch_Click_Err")

End Sub

Private Sub cmdSearch2_Click()

    Dim strSQL As String
    Dim strWhere As String

    On Error GoTo cmdSearch2_Click_Err

    Screen.MousePointer = vbHourglass

    cmdSearch2.SetFocus

    strSQL = "SELECT * FROM DT021 "
    
    strWhere = " LEFT(Ocode,8) = '" & Global_Get_NumericDay(frmYpmf020.lblOdate.Caption) & "'"
    If Trim(txtPnum.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " Pnum = '" & txtPnum.Text & "'"
    End If
    If Trim(txtBcode.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " Bcode = '" & txtBcode.Text & "'"
    End If
    
    '2005/09/01 追加
    If Trim(txtIname.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " Iname LIKE '%" & txtIname.Text & "%'"
    End If
    
    '検索実行
    If Trim(strWhere) <> "" Then
        Adodc2.RecordSource = strSQL & " WHERE " & strWhere & " ORDER BY Pnum,PnumLine"
    Else
        Adodc2.RecordSource = strSQL & " ORDER BY Pnum,PnumLine"
    End If
    Adodc2.Refresh
    Set DataGrid2.DataSource = Adodc2

    If Adodc2.Recordset.BOF Or Adodc2.Recordset.EOF Then
        DataGrid2.Visible = False
        lblDummy2.Visible = True
        Call MsgBox("データがありません！！", vbOKOnly + vbInformation, "情報")
    Else
        DataGrid2.Visible = True
        lblDummy2.Visible = False
    End If

    Screen.MousePointer = vbDefault

    Exit Sub

cmdSearch2_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("検索開始クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdSearch2_Click_Err")

End Sub

Private Sub DataGrid1_DblClick()

    Call cmdExecute_Click
    
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)

    m_intHeadClickCol = ColIndex
    PopupMenu mnuPop

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error Resume Next
    
    txtKey1.Text = IIf(IsNull(DataGrid1.Columns(0)), "", Trim(DataGrid1.Columns(0)))

End Sub

Private Sub DataGrid2_DblClick()

    Call cmdExecute_Click
    
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)

    m_intHeadClickCol = ColIndex
    PopupMenu mnuPop

End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error Resume Next
    
    txtKey1.Text = IIf(IsNull(DataGrid2.Columns(5)), "", Trim(DataGrid2.Columns(5)))

End Sub

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
        Case vbKeyF9
        Case vbKeyF10
        Case vbKeyF11
        Case vbKeyF12
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

    On Error Resume Next

    txtOcode.Text = ""
    txtHnum.Text = ""
    txtPnum.Text = ""
    txtBcode.Text = ""
    txtIname.Text = ""
    
    txtKey1.Text = ""
    DataGrid1.Visible = False
    lblDummy.Visible = True
    DataGrid2.Visible = False
    lblDummy2.Visible = True

    fraSearch1.Visible = True
    fraSearch2.Visible = False

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    If TabStrip1.Tabs(1).Selected = True Then
        txtOcode.SetFocus
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        txtPnum.SetFocus
    End If

End Sub

Private Sub TabStrip1_Click()

    If TabStrip1.Tabs(1).Selected = True Then
        fraSearch1.Visible = True
        fraSearch2.Visible = False
        cmdPrint.Visible = False
        txtOcode.SetFocus
        DoEvents
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        fraSearch1.Visible = False
        fraSearch2.Visible = True
        cmdPrint.Visible = True
        txtPnum.SetFocus
        DoEvents
    End If

End Sub

Private Sub txtBcode_GotFocus()
    
    txtBcode.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtBcode_LostFocus()
    
    txtBcode.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtIname_GotFocus()
    
    txtIname.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtIname_LostFocus()
   
    txtIname.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtOcode_GotFocus()
    
    txtOcode.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtOcode_LostFocus()
   
    txtOcode.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtHnum_GotFocus()
    
    txtHnum.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtHnum_LostFocus()
    
    txtHnum.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub Order_Data(ColIndex As Integer, intFlg As Integer)
   
    Dim strSQL As String
    Dim strOrder As String
    
    On Error GoTo Order_Data_Err
    
    Screen.MousePointer = vbHourglass
    
    'ソート条件
    strOrder = " ORDER BY "
    Select Case ColIndex
        Case 0:
            strOrder = strOrder & "Ocode"
        Case 1:
            strOrder = strOrder & "Hnum"
    End Select
    
    If intFlg = 0 Then
        '昇順
        strOrder = strOrder & " ASC"
    ElseIf intFlg = 1 Then
        '降順
        strOrder = strOrder & " DESC"
    End If
    
    'ORDER BY句の抜き出し
    If InStrRev(Adodc1.RecordSource, "ORDER BY") > 0 Then
        strSQL = left(Adodc1.RecordSource, InStrRev(Adodc1.RecordSource, "ORDER BY") - 1)
    Else
        strSQL = Adodc1.RecordSource
    End If
    
    Adodc1.RecordSource = strSQL & strOrder
    Adodc1.Refresh
    
    If Adodc1.Recordset.BOF = True Or Adodc1.Recordset.EOF = True Then
        Call MsgBox("データがありません！！", vbOKOnly + vbInformation, "情報")
    End If
    DataGrid1.Refresh
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Order_Data_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("並び替え時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Order_Data_Err")
                
End Sub

Private Sub mnuAsc_Click()
    
    If TabStrip1.Tabs(1).Selected = True Then
        Call Order_Data(m_intHeadClickCol, 0)
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        Call Order_Data2(m_intHeadClickCol, 0)
    End If

End Sub

Private Sub mnuDesc_Click()
    
    If TabStrip1.Tabs(1).Selected = True Then
        Call Order_Data(m_intHeadClickCol, 1)
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        Call Order_Data2(m_intHeadClickCol, 1)
    End If
   
End Sub

Private Sub txtPnum_GotFocus()
    
    txtPnum.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtPnum_LostFocus()
    
    txtPnum.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub Order_Data2(ColIndex As Integer, intFlg As Integer)
   
    Dim strSQL As String
    Dim strOrder As String
    
    On Error GoTo Order_Data2_Err
    
    Screen.MousePointer = vbHourglass
    
    'ソート条件
    strOrder = " ORDER BY "
    Select Case ColIndex
        Case 0:
            strOrder = strOrder & "Pnum"
        Case 1:
            strOrder = strOrder & "PnumLine"
        Case 2:
            strOrder = strOrder & "Iname"
        Case 3:
            strOrder = strOrder & "Price"
        Case 4:
            strOrder = strOrder & "Bcode"
        Case 5:
            strOrder = strOrder & "Ocode"
    End Select
    
    If intFlg = 0 Then
        '昇順
        strOrder = strOrder & " ASC"
    ElseIf intFlg = 1 Then
        '降順
        strOrder = strOrder & " DESC"
    End If
    
    'ORDER BY句の抜き出し
    If InStrRev(Adodc2.RecordSource, "ORDER BY") > 0 Then
        strSQL = left(Adodc2.RecordSource, InStrRev(Adodc2.RecordSource, "ORDER BY") - 1)
    Else
        strSQL = Adodc2.RecordSource
    End If
    
    Adodc2.RecordSource = strSQL & strOrder
    Adodc2.Refresh
    
    If Adodc2.Recordset.BOF = True Or Adodc2.Recordset.EOF = True Then
        Call MsgBox("データがありません！！", vbOKOnly + vbInformation, "情報")
    End If
    DataGrid2.Refresh
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Order_Data2_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("並び替え時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Order_Data2_Err")
                
End Sub


