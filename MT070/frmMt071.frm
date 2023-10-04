VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmMt071 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "買主番号履歴"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "frmMt071.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdClear_Dst 
      Caption         =   "明細クリア"
      Height          =   375
      Left            =   3900
      TabIndex        =   10
      Top             =   1980
      Width           =   1575
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "明細登録"
      Height          =   375
      Left            =   780
      TabIndex        =   8
      Top             =   1980
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "明細削除"
      Height          =   375
      Left            =   2340
      TabIndex        =   9
      Top             =   1980
      Width           =   1575
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
      Height          =   375
      Left            =   2700
      TabIndex        =   12
      Top             =   4140
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      TabIndex        =   13
      Top             =   4140
      Width           =   1335
   End
   Begin MSComctlLib.ListView lsvMeisai 
      Height          =   1635
      Left            =   60
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2400
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2884
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "行"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "開始年月日"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "終了年月日"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "買主ｺｰﾄﾞ"
         Object.Width           =   2540
      EndProperty
   End
   Begin imText6Ctl.imText txtYear 
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "frmMt071.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":007A
      Key             =   "frmMt071.frx":0098
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
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   2
      Top             =   60
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   635
      Caption         =   "frmMt071.frx":00CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":013A
      Key             =   "frmMt071.frx":0158
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
      Height          =   360
      Index           =   0
      Left            =   3660
      TabIndex        =   3
      Top             =   60
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   635
      Caption         =   "frmMt071.frx":018C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":01FA
      Key             =   "frmMt071.frx":0218
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
   Begin imText6Ctl.imText txtYear 
      Height          =   360
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "frmMt071.frx":024C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":02BA
      Key             =   "frmMt071.frx":02D8
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
      Height          =   360
      Index           =   1
      Left            =   2700
      TabIndex        =   5
      Top             =   840
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   635
      Caption         =   "frmMt071.frx":030C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":037A
      Key             =   "frmMt071.frx":0398
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
      Height          =   360
      Index           =   1
      Left            =   3660
      TabIndex        =   6
      Top             =   840
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   635
      Caption         =   "frmMt071.frx":03CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":043A
      Key             =   "frmMt071.frx":0458
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
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   5
      Left            =   60
      TabIndex        =   22
      Top             =   60
      Width           =   1455
      _Version        =   262145
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "日付範囲"
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
      Left            =   60
      TabIndex        =   23
      Top             =   1380
      Width           =   1455
      _Version        =   262145
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "買主コード"
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
   Begin imText6Ctl.imText txtBnum 
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   1380
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "frmMt071.frx":048C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":04FA
      Key             =   "frmMt071.frx":0518
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
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   5880
      TabIndex        =   0
      Top             =   0
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmMt071.frx":054C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":05BA
      Key             =   "frmMt071.frx":05D8
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
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmMt071.frx":061C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":068A
      Key             =   "frmMt071.frx":06A8
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
   Begin imNumber6Ctl.imNumber imnNo 
      Height          =   375
      Left            =   60
      TabIndex        =   24
      Top             =   1980
      Visible         =   0   'False
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   661
      Calculator      =   "frmMt071.frx":06EC
      Caption         =   "frmMt071.frx":070C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":077A
      Keys            =   "frmMt071.frx":0798
      Spin            =   "frmMt071.frx":07E2
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
      Value           =   1
      MaxValueVT      =   1230438405
      MinValueVT      =   1313734661
   End
   Begin imText6Ctl.imText txtBcode 
      Height          =   360
      Left            =   60
      TabIndex        =   25
      Top             =   4140
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "frmMt071.frx":080A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt071.frx":0878
      Key             =   "frmMt071.frx":0896
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
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
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
      Index           =   6
      Left            =   2760
      TabIndex        =   21
      Top             =   480
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "日"
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
      Index           =   5
      Left            =   4200
      TabIndex        =   20
      Top             =   900
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "月"
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
      Index           =   4
      Left            =   3240
      TabIndex        =   19
      Top             =   900
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "年"
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
      Index           =   3
      Left            =   2340
      TabIndex        =   18
      Top             =   900
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "日"
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
      Index           =   2
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "月"
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
      Index           =   1
      Left            =   3240
      TabIndex        =   16
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "年"
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
      Index           =   0
      Left            =   2340
      TabIndex        =   15
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "frmMt071"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdClear_Dst_Click()

    Call FieldsClear(1)
    Call ListViewGetMaxRow
    txtYear(0).SetFocus
    
End Sub

Private Sub cmdDel_Click()

    If ListViewDelItem() = False Then Exit Sub
    Call FieldsClear(1)
    txtYear(0).SetFocus

End Sub

Private Sub cmdEdit_Click()

    If DoValidationChecks_Dst() = False Then Exit Sub
    If ListViewSetItem(imnNo.Value, 0) = False Then Exit Sub
    Call FieldsClear(1)
    txtYear(0).SetFocus
    
End Sub

Private Sub cmdExecute_Click()

    If DoValidationChecks() = False Then Exit Sub
    If DataUpdate() = True Then Unload Me

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

    Call FieldsClear(0)
    txtBcode.Text = frmMt070.txtBcode
    Call FieldsSet

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdCancel.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    txtYear(0).SetFocus
    
End Sub

Private Sub lsvMeisai_Click()

    On Error Resume Next

    '行が選択されているか？
    If lsvMeisai.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    '明細表示
    Call ListViewGetItem
    
    txtYear(0).SetFocus

End Sub

Private Sub txtBnum_GotFocus()
    
    txtBnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtBnum_LostFocus()
    
    txtBnum.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtDay_GotFocus(Index As Integer)
    
    txtDay(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtDay_LostFocus(Index As Integer)
    
    txtDay(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtMonth_GotFocus(Index As Integer)
    
    txtMonth(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtMonth_LostFocus(Index As Integer)
    
    txtMonth(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtYear_GotFocus(Index As Integer)
    
    txtYear(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtYear_LostFocus(Index As Integer)
    
    txtYear(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：リストビューへのデータ登録
'条　件　　：
'結　果　　：
'引　数　　：intFlg(0:追加・更新 1:挿入)
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
'更新履歴　：
'
Private Function ListViewSetItem(intPostion As Integer, intFlg As Integer) As Boolean

    Dim itmX As ListItem

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
    itmX.SubItems(1) = Trim(txtYear(0).Text) & "/" & Format(txtMonth(0).Text, "00") & "/" & Format(txtDay(0).Text, "00")
    itmX.SubItems(2) = Trim(txtYear(1).Text) & "/" & Format(txtMonth(1).Text, "00") & "/" & Format(txtDay(1).Text, "00")
    itmX.SubItems(3) = Trim(txtBnum.Text)

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

'目　的　　：リストビューからの行番号取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
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

'目　的　　：リストビューからのデータ表示
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
'更新履歴　：
'
Private Sub ListViewGetItem()

    On Error GoTo ListViewGetItem_Err
    
    imnNo.Value = lsvMeisai.SelectedItem.Text
    txtYear(0).Text = Mid(lsvMeisai.SelectedItem.SubItems(1), 1, 4)
    txtMonth(0).Text = Mid(lsvMeisai.SelectedItem.SubItems(1), 6, 2)
    txtDay(0).Text = Mid(lsvMeisai.SelectedItem.SubItems(1), 9, 2)
    txtYear(1).Text = Mid(lsvMeisai.SelectedItem.SubItems(2), 1, 4)
    txtMonth(1).Text = Mid(lsvMeisai.SelectedItem.SubItems(2), 6, 2)
    txtDay(1).Text = Mid(lsvMeisai.SelectedItem.SubItems(2), 9, 2)
    txtBnum.Text = lsvMeisai.SelectedItem.SubItems(3)
        
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
'作成年月日：２００２／０６／２１
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

'目　的　　：明細入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
'更新履歴　：
'
Private Function DoValidationChecks_Dst() As Boolean

    Dim strErrMsg As String
    Dim strDate As String
        
    On Error GoTo DoValidationChecks_Dst_Err

    If Trim(txtYear(0).Text) = "" Or Trim(txtMonth(0).Text) = "" Or Trim(txtDay(0).Text) = "" Then
        strErrMsg = "日付を入力してください。"
        txtYear(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtYear(1).Text) = "" Or Trim(txtMonth(1).Text) = "" Or Trim(txtDay(1).Text) = "" Then
        strErrMsg = "日付を入力してください。"
        txtYear(1).SetFocus
        GoTo ErrorTrap:
    End If
    If CLng(txtYear(0).Text) < 1900 Then
        strErrMsg = "西暦年４桁で入力してください。"
        txtYear(0).SetFocus
        GoTo ErrorTrap:
    End If
    If CLng(txtYear(1).Text) < 1900 Then
        strErrMsg = "西暦年４桁で入力してください。"
        txtYear(1).SetFocus
        GoTo ErrorTrap:
    End If
    If Global_IsDate(Trim(txtYear(0).Text), Trim(txtMonth(0).Text), Trim(txtDay(0).Text)) = False Then
        strErrMsg = "正しい日付を入力してください。"
        txtYear(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Global_IsDate(Trim(txtYear(1).Text), Trim(txtMonth(1).Text), Trim(txtDay(1).Text)) = False Then
        strErrMsg = "正しい日付を入力してください。"
        txtYear(1).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtBnum.Text) = "" Then
        strErrMsg = "買主コードを入力してください。"
        txtBnum.SetFocus
        GoTo ErrorTrap:
    End If
    If CheckBcode() = False Then
        strErrMsg = "買主コードが重複しています。"
        txtBnum.SetFocus
        GoTo ErrorTrap:
    End If
    If CheckSdate() = False Then
        strErrMsg = "開始年月日が重複しています。"
        txtYear(0).SetFocus
        GoTo ErrorTrap:
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

'目　的　　：リストビューの行番号を振り直す
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
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

'目　的　　：画面クリア
'条　件　　：
'結　果　　：
'引　数　　：0：全画面 1:明細部
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        txtYear(0).Text = ""
        txtMonth(0).Text = ""
        txtDay(0).Text = ""
        txtYear(1).Text = ""
        txtMonth(1).Text = ""
        txtDay(1).Text = ""
        txtBnum.Text = ""
        imnNo.Value = 1
        lsvMeisai.ListItems.Clear
        txtBcode.Text = ""
    ElseIf intKubun = 1 Then
        txtYear(0).Text = ""
        txtMonth(0).Text = ""
        txtDay(0).Text = ""
        txtYear(1).Text = ""
        txtMonth(1).Text = ""
        txtDay(1).Text = ""
        txtBnum.Text = ""
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

Private Function CheckBcode() As Boolean

    Dim intIndex1 As Integer

    On Error GoTo CheckBcode_Err
    
    CheckBcode = True
    
    lsvMeisai.Refresh
    For intIndex1 = 1 To lsvMeisai.ListItems.Count Step 1
        If imnNo.Value <> intIndex1 Then
            If Trim(txtBnum.Text) = Trim(lsvMeisai.ListItems(intIndex1).SubItems(3)) Then
                CheckBcode = False
            End If
        End If
    Next intIndex1
    
    Exit Function

CheckBcode_Err:

    CheckBcode = False
    Screen.MousePointer = vbDefault
    Call MsgBox("買主コードのチェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "CheckBcode_Err")

End Function

Private Function CheckSdate() As Boolean

    Dim intIndex1 As Integer

    On Error GoTo CheckSdate_Err
    
    CheckSdate = True
    
    lsvMeisai.Refresh
    For intIndex1 = 1 To lsvMeisai.ListItems.Count Step 1
        If imnNo.Value <> intIndex1 Then
            If Global_StrToDate(txtYear(0).Text, txtMonth(0).Text, txtDay(0).Text) = Trim(lsvMeisai.ListItems(intIndex1).SubItems(1)) Then
                CheckSdate = False
            End If
        End If
    Next intIndex1
    
    Exit Function

CheckSdate_Err:

    CheckSdate = False
    Screen.MousePointer = vbDefault
    Call MsgBox("開始年月日のチェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "CheckSdate_Err")

End Function

'目　的　　：データの登録
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.RecordSet
    Dim intIndex1 As Integer

    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    frmMt070.m_clsAdoSQL.Connection.BeginTrans
 
    'データ削除
    strSQL = "DELETE FROM MT071" & _
             " WHERE Bcode = " & txtBcode.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    With adoRecordset1
        strSQL = "SELECT * FROM MT071" & _
                 " WHERE Bcode = " & txtBcode.Text
        .Open strSQL, frmMt070.m_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            .AddNew
            .Fields("Bcode") = txtBcode.Text
            .Fields("Sdate") = lsvMeisai.ListItems(intIndex1).SubItems(1)
            .Fields("Fdate") = lsvMeisai.ListItems(intIndex1).SubItems(2)
            .Fields("Bnum") = lsvMeisai.ListItems(intIndex1).SubItems(3)
            .Update
        Next intIndex1
        .Close
    End With
    
    frmMt070.m_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    Set adoRecordset1 = Nothing
    
    DataUpdate = True
    
    Exit Function

DataUpdate_Err:

    frmMt070.m_clsAdoSQL.Connection.RollbackTrans
    DataUpdate = False
    Screen.MousePointer = vbDefault
    Call MsgBox("データ登録エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_Err")

End Function

'目　的　　：フィールドのセット
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１３
'更新履歴　：
'
Public Function FieldsSet() As Boolean
    
    Dim adoRecordset1 As New ADODB.RecordSet
    Dim strSQL As String
    Dim itmX As ListItem
    Dim intIndex1 As Integer

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    With adoRecordset1
        strSQL = "SELECT * FROM MT071" & _
                 " WHERE Bcode = " & txtBcode.Text
        .Open strSQL, frmMt070.m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        
        intIndex1 = 1
        lsvMeisai.ListItems.Clear
        Do While Not .EOF
            Set itmX = lsvMeisai.ListItems.Add(, , intIndex1, 0)
            
            itmX.SubItems(1) = IIf(IsNull(.Fields("Sdate")), "", .Fields("Sdate"))
            itmX.SubItems(2) = IIf(IsNull(.Fields("Fdate")), "", .Fields("Fdate"))
            itmX.SubItems(3) = IIf(IsNull(.Fields("Bnum")), "", .Fields("Bnum"))
            
            intIndex1 = intIndex1 + 1
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
        
        Call ListViewGetMaxRow
    End With

    Screen.MousePointer = vbDefault
    
    FieldsSet = True
    
    Exit Function

FieldsSet_Err:

    FieldsSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("フィールドセットエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsSet_Err")

End Function

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

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If lsvMeisai.ListItems.Count <= 0 Then
        strErrMsg = "明細を入力してください。"
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

