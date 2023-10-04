VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUekiSearch 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "êAñÿåüçı"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "frmUekiSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.TextBox txtKey1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   3
      Top             =   7620
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "ñﬂÇÈ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8460
      TabIndex        =   1
      Top             =   7560
      Width           =   1875
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
      Caption         =   "frmUekiSearch.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUekiSearch.frx":007A
      Key             =   "frmUekiSearch.frx":0098
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
      TabIndex        =   2
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmUekiSearch.frx":00DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUekiSearch.frx":014A
      Key             =   "frmUekiSearch.frx":0168
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
      Left            =   60
      Top             =   7620
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
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fraSearch2 
      Height          =   7455
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   10275
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmUekiSearch.frx":01AC
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
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            Caption         =   "éÛïtî‘çÜ"
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
            DataField       =   "Line"
            Caption         =   "çs"
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
            Caption         =   "êAñÿñºèÃ"
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
            Caption         =   "îÑóßã‡äz"
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
            Caption         =   "îÉéÂ"
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
            Caption         =   "ã£îÑî‘çÜ"
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
         Caption         =   "åüçıèåè"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         TabIndex        =   5
         Top             =   180
         Width           =   10035
         Begin VB.CommandButton cmdSearch2 
            Caption         =   "åüçıäJén"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   8040
            TabIndex        =   10
            Top             =   240
            Width           =   1875
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   300
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            TabIndex        =   7
            Top             =   300
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
            _ExtentY        =   635
            Caption         =   "frmUekiSearch.frx":01C1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmUekiSearch.frx":022F
            Key             =   "frmUekiSearch.frx":024D
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
            TabIndex        =   13
            Top             =   300
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "îÉéÂî‘çÜ"
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
            TabIndex        =   8
            Top             =   300
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
            _ExtentY        =   635
            Caption         =   "frmUekiSearch.frx":0291
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmUekiSearch.frx":02FF
            Key             =   "frmUekiSearch.frx":031D
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
            TabIndex        =   14
            Top             =   300
            Width           =   855
            _Version        =   262145
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "êAñÿñº"
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
            TabIndex        =   9
            Top             =   300
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   609
            Caption         =   "frmUekiSearch.frx":0361
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmUekiSearch.frx":03CF
            Key             =   "frmUekiSearch.frx":03ED
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
         Appearance      =   0  'Ã◊Øƒ
         BorderStyle     =   1  'é¿ê¸
         ForeColor       =   &H80000008&
         Height          =   6315
         Left            =   120
         TabIndex        =   12
         Top             =   1020
         Width           =   10035
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "É|ÉbÉvÉAÉbÉvÉÅÉjÉÖÅ["
      Visible         =   0   'False
      Begin VB.Menu mnuAsc 
         Caption         =   "è∏Å@èá"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesc 
         Caption         =   "ç~Å@èá"
      End
   End
End
Attribute VB_Name = "frmUekiSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_intHeadClickCol As Integer            'ÉwÉbÉ_Å[ÉNÉäÉbÉNóÒ

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdSearch2_Click()

    Dim strSQL As String
    Dim strWhere As String

    On Error GoTo cmdSearch2_Click_Err

    Screen.MousePointer = vbHourglass

    cmdSearch2.SetFocus

    strSQL = "SELECT DT011.Pnum,DT011.Line,DT011.Iname,vw_YPMF021.Price,vw_YPMF021.Bcode,vw_YPMF021.Ocode FROM DT011 "
    strSQL = strSQL & " LEFT OUTER JOIN vw_YPMF021 ON vw_YPMF021.Odate = DT011.Odate AND vw_YPMF021.Pnum = DT011.Pnum AND vw_YPMF021.PnumLine = DT011.Line "
    
    strWhere = " DT011.Odate = '" & frmYpmf020.lblOdate.Caption & "'"

    If Trim(txtPnum.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " DT011.Pnum = " & txtPnum.Text
    End If
    If Trim(txtBcode.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " vw_YPMF021.Bcode = '" & txtBcode.Text & "'"
    End If
    If Trim(txtIname.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " DT011.Iname LIKE '%" & txtIname.Text & "%'"
    End If
    
    'åüçıé¿çs
    If Trim(strWhere) <> "" Then
        Adodc2.RecordSource = strSQL & " WHERE " & strWhere & " ORDER BY DT011.Pnum,DT011.Line"
    Else
        Adodc2.RecordSource = strSQL & " ORDER BY DT011.Pnum,DT011.Line"
    End If
    Adodc2.Refresh
    Set DataGrid2.DataSource = Adodc2

    If Adodc2.Recordset.BOF Or Adodc2.Recordset.EOF Then
        DataGrid2.Visible = False
        lblDummy2.Visible = True
        Call MsgBox("ÉfÅ[É^Ç™Ç†ÇËÇ‹ÇπÇÒÅIÅI", vbOKOnly + vbInformation, "èÓïÒ")
    Else
        DataGrid2.Visible = True
        lblDummy2.Visible = False
    End If

    Screen.MousePointer = vbDefault

    Exit Sub

cmdSearch2_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("åüçıäJénÉNÉäÉbÉNéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdSearch2_Click_Err")

End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)

    m_intHeadClickCol = ColIndex
    PopupMenu mnuPop

End Sub

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

    Call MsgBox("ÉtÉHÅ[ÉÄÉLÅ[É_ÉEÉìéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")

End Sub

Private Sub Form_Load()

    On Error Resume Next

    txtPnum.Text = ""
    txtBcode.Text = ""
    txtIname.Text = ""
    
    txtKey1.Text = ""
    DataGrid2.Visible = False
    lblDummy2.Visible = True

    fraSearch2.Visible = True

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    txtPnum.SetFocus

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

Private Sub mnuAsc_Click()
    
    Call Order_Data2(m_intHeadClickCol, 0)

End Sub

Private Sub mnuDesc_Click()
    
    Call Order_Data2(m_intHeadClickCol, 1)
   
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
    
    'É\Å[Égèåè
    strOrder = " ORDER BY "
    Select Case ColIndex
        Case 0:
            strOrder = strOrder & "DT011.Pnum"
        Case 1:
            strOrder = strOrder & "DT011.Line"
        Case 2:
            strOrder = strOrder & "DT011.Iname"
        Case 3:
            strOrder = strOrder & "vw_YPMF021.Price"
        Case 4:
            strOrder = strOrder & "vw_YPMF021.Bcode"
        Case 5:
            strOrder = strOrder & "vw_YPMF021.Ocode"
    End Select
    
    If intFlg = 0 Then
        'è∏èá
        strOrder = strOrder & " ASC"
    ElseIf intFlg = 1 Then
        'ç~èá
        strOrder = strOrder & " DESC"
    End If
    
    'ORDER BYãÂÇÃî≤Ç´èoÇµ
    If InStrRev(Adodc2.RecordSource, "ORDER BY") > 0 Then
        strSQL = left(Adodc2.RecordSource, InStrRev(Adodc2.RecordSource, "ORDER BY") - 1)
    Else
        strSQL = Adodc2.RecordSource
    End If
    
    Adodc2.RecordSource = strSQL & strOrder
    Adodc2.Refresh
    
    If Adodc2.Recordset.BOF = True Or Adodc2.Recordset.EOF = True Then
        Call MsgBox("ÉfÅ[É^Ç™Ç†ÇËÇ‹ÇπÇÒÅIÅI", vbOKOnly + vbInformation, "èÓïÒ")
    End If
    DataGrid2.Refresh
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Order_Data2_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("ï¿Ç—ë÷Ç¶éûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Order_Data2_Err")
                
End Sub


