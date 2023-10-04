VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSearch2 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "出品者受付検索"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13275
   Icon            =   "frmSearch2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSearch2.frx":000C
      Height          =   6855
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   12091
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Pnum"
         Caption         =   "番号"
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
         DataField       =   "Scode"
         Caption         =   "ｺｰﾄﾞ"
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
         DataField       =   "Sname"
         Caption         =   "出品者名"
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
         DataField       =   "Addres"
         Caption         =   "住　所"
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
      BeginProperty Column04 
         DataField       =   "Tel"
         Caption         =   "電話番号"
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
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4229.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   5504.882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1934.929
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
      Height          =   1395
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   13035
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
         Left            =   11040
         TabIndex        =   4
         Top             =   300
         Width           =   1875
      End
      Begin imText6Ctl.imText txtSname 
         Height          =   345
         Left            =   1620
         TabIndex        =   1
         Top             =   300
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   609
         Caption         =   "frmSearch2.frx":0021
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSearch2.frx":008F
         Key             =   "frmSearch2.frx":00AD
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   300
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
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "住　所"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   5100
         TabIndex        =   14
         Top             =   300
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
      Begin imText6Ctl.imText txtAddres 
         Height          =   585
         Left            =   1620
         TabIndex        =   3
         Top             =   720
         Width           =   8595
         _Version        =   65536
         _ExtentX        =   15161
         _ExtentY        =   1032
         Caption         =   "frmSearch2.frx":00F1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSearch2.frx":015F
         Key             =   "frmSearch2.frx":017D
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
      Begin imText6Ctl.imText txtTel 
         Height          =   360
         Left            =   6600
         TabIndex        =   2
         Top             =   300
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   635
         Caption         =   "frmSearch2.frx":01C1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSearch2.frx":022F
         Key             =   "frmSearch2.frx":024D
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
      Left            =   2340
      TabIndex        =   9
      Top             =   8520
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
      Left            =   11280
      TabIndex        =   7
      Top             =   8520
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
      Left            =   9300
      TabIndex        =   6
      Top             =   8520
      Width           =   1875
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   120
      Top             =   8520
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "SELECT * FROM DT010 ORDER BY Odate,Pnum"
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
      Left            =   13380
      TabIndex        =   0
      Top             =   300
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmSearch2.frx":0281
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearch2.frx":02EF
      Key             =   "frmSearch2.frx":030D
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
      Left            =   13620
      TabIndex        =   8
      Top             =   300
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmSearch2.frx":0351
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearch2.frx":03BF
      Key             =   "frmSearch2.frx":03DD
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
   Begin VB.Label lblDummy 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   13035
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
Attribute VB_Name = "frmSearch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_intHeadClickCol As Integer            'ヘッダークリック列

Private Sub cmdExecute_Click()

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cmdExecute_Click_Err

    Screen.MousePointer = vbHourglass

    If Trim(txtKey1.Text) = "" Then
        Call MsgBox("データを選択してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    
    frmYpmf030.txtPnum.Text = txtKey1.Text
    
    '受付データ
    strSQL = "SELECT * FROM DT010" & _
             " WHERE Odate = '" & g_strOdate & "'" & _
             " AND Pnum = " & txtKey1.Text & _
             " ORDER BY Odate,Pnum"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        frmYpmf030.txtSname.Text = IIf(IsNull(adoRecordset1.Fields("Sname")), "", Trim(adoRecordset1.Fields("Sname")))
        frmYpmf030.txtAddres.Text = IIf(IsNull(adoRecordset1.Fields("Addres")), "", Trim(adoRecordset1.Fields("Addres")))
        frmYpmf030.txtTel.Text = IIf(IsNull(adoRecordset1.Fields("Tel")), "", Trim(adoRecordset1.Fields("Tel")))
        If Not IsNull(adoRecordset1.Fields("Div")) Then
            If adoRecordset1.Fields("Div") = TIKU_DIV_ON Then
                frmYpmf030.optDiv(0).Value = True
            Else
                frmYpmf030.optDiv(1).Value = True
            End If
        Else
            frmYpmf030.optDiv(0).Value = True
        End If
        frmYpmf030.chkSoukin.Value = IIf(IsNull(adoRecordset1.Fields("Soukin")), 0, adoRecordset1.Fields("Soukin"))
    End If
    adoRecordset1.Close
    
    Screen.MousePointer = vbDefault
    Unload Me
    
    Exit Sub
    
cmdExecute_Click_Err:
    
    Screen.MousePointer = vbDefault
    Call MsgBox("ＯＫクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")
    
End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdSearch_Click()

    Dim strSQL As String
    Dim strWhere As String

    On Error GoTo cmdSearch_Click_Err

    Screen.MousePointer = vbHourglass

    cmdSearch.SetFocus

    strSQL = "SELECT * FROM DT010 "
    strWhere = " Odate = '" & frmYpmf030.lblOdate.Caption & "'"

    If Trim(txtSname.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " Sname LIKE '%" & txtSname.Text & "%'"
    End If
    If Trim(txtTel.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " Tel LIKE '%" & txtTel.Text & "%'"
    End If
    If Trim(txtAddres.Text) <> "" Then
        If Trim(strWhere) <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " Addres LIKE '%" & txtAddres.Text & "%'"
    End If

    '検索実行
    If Trim(strWhere) <> "" Then
        Adodc1.RecordSource = strSQL & " WHERE " & strWhere
    Else
        Adodc1.RecordSource = strSQL
    End If
    Adodc1.Refresh

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

    txtSname.Text = ""
    txtTel.Text = ""
    txtAddres.Text = ""
    txtKey1.Text = ""

    DataGrid1.Visible = False
    lblDummy.Visible = True

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    txtSname.SetFocus

End Sub

Private Sub txtAddres_GotFocus()
    
    txtAddres.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtAddres_LostFocus()
   
    txtAddres.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtSname_GotFocus()
    
    txtSname.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtSname_LostFocus()
    
    txtSname.BackColor = FOCUS_NO_COLOR

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
            strOrder = strOrder & "Pnum"
        Case 1:
            strOrder = strOrder & "Scode"
        Case 2:
            strOrder = strOrder & "Sname"
        Case 3:
            strOrder = strOrder & "Addres"
        Case 4:
            strOrder = strOrder & "Tel"
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

    Call Order_Data(m_intHeadClickCol, 0)

End Sub

Private Sub mnuDesc_Click()

   Call Order_Data(m_intHeadClickCol, 1)
   
End Sub

Private Sub txtTel_GotFocus()
    
    txtTel.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtTel_LostFocus()

    txtTel.BackColor = FOCUS_NO_COLOR
    
End Sub
