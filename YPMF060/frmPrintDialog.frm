VERSION 5.00
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmPrintDialog 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "印刷"
   ClientHeight    =   1515
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8790
   Icon            =   "frmPrintDialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Caption         =   "抽出条件"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   7275
      Begin VB.OptionButton optOrder 
         Caption         =   "買主コード順"
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
         Left            =   4620
         TabIndex        =   15
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "入金種別、入金日順"
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
         Left            =   1680
         TabIndex        =   14
         Top             =   840
         Value           =   -1  'True
         Width           =   2595
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   5
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "入金日"
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
      Begin imNumber6Ctl.imNumber imnRdate_Year 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   661
         Calculator      =   "frmPrintDialog.frx":000C
         Caption         =   "frmPrintDialog.frx":002C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":009A
         Keys            =   "frmPrintDialog.frx":00B8
         Spin            =   "frmPrintDialog.frx":00F2
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
      Begin imNumber6Ctl.imNumber imnRdate_Month 
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmPrintDialog.frx":011A
         Caption         =   "frmPrintDialog.frx":013A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":01A8
         Keys            =   "frmPrintDialog.frx":01C6
         Spin            =   "frmPrintDialog.frx":0200
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "00"
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
      Begin imNumber6Ctl.imNumber imnRdate_Day 
         Height          =   375
         Index           =   0
         Left            =   3420
         TabIndex        =   7
         Top             =   360
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmPrintDialog.frx":0228
         Caption         =   "frmPrintDialog.frx":0248
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":02B6
         Keys            =   "frmPrintDialog.frx":02D4
         Spin            =   "frmPrintDialog.frx":030E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "00"
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
         ValueVT         =   2012217349
         Value           =   31
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imNumber6Ctl.imNumber imnRdate_Year 
         Height          =   375
         Index           =   1
         Left            =   4620
         TabIndex        =   11
         Top             =   360
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   661
         Calculator      =   "frmPrintDialog.frx":0336
         Caption         =   "frmPrintDialog.frx":0356
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":03C4
         Keys            =   "frmPrintDialog.frx":03E2
         Spin            =   "frmPrintDialog.frx":041C
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
      Begin imNumber6Ctl.imNumber imnRdate_Month 
         Height          =   375
         Index           =   1
         Left            =   5580
         TabIndex        =   12
         Top             =   360
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmPrintDialog.frx":0444
         Caption         =   "frmPrintDialog.frx":0464
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":04D2
         Keys            =   "frmPrintDialog.frx":04F0
         Spin            =   "frmPrintDialog.frx":052A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "00"
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
      Begin imNumber6Ctl.imNumber imnRdate_Day 
         Height          =   375
         Index           =   1
         Left            =   6360
         TabIndex        =   13
         Top             =   360
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmPrintDialog.frx":0552
         Caption         =   "frmPrintDialog.frx":0572
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":05E0
         Keys            =   "frmPrintDialog.frx":05FE
         Spin            =   "frmPrintDialog.frx":0638
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "00"
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
         ValueVT         =   1245189
         Value           =   31
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   840
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "出力順"
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
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   18
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   6060
         TabIndex        =   17
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   6840
         TabIndex        =   16
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   30
         Left            =   2340
         TabIndex        =   10
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   31
         Left            =   3120
         TabIndex        =   9
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   32
         Left            =   3900
         TabIndex        =   8
         Top             =   420
         Width           =   315
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
         Index           =   0
         Left            =   4260
         TabIndex        =   4
         Top             =   420
         Width           =   375
      End
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
      Left            =   7440
      TabIndex        =   1
      Top             =   660
      Width           =   1275
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
      Left            =   7440
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objArPrint As New clsArPrint
Public m_strRdateFrom As String
Public m_strRdateTo As String

'2005/08/12 修正
Private Sub cmdExecute_Click()
    
    Dim objRpt1 As New rptYpmf060
    Dim objRpt2 As New rptYpmf060_2
    
    On Error GoTo cmdExecute_Click_Err

    If DoValidationChecks() = False Then Exit Sub
    If MakeWork() = False Then Exit Sub

    cmdCancel.SetFocus
    DoEvents

    Screen.MousePointer = vbHourglass

    With objArPrint
        .Name = "入金一覧表"
        If optOrder(0).Value = True Then
            .objReport = objRpt2
        ElseIf optOrder(1).Value = True Then
            .objReport = objRpt1
        End If
        .Connection = g_clsAdoAccess.Connection
        
        If optOrder(0).Value = True Then
            .SQL = "SELECT * FROM QWK_YPMF060 ORDER BY R,Rdate,Bcode,Odate"
        ElseIf optOrder(1).Value = True Then
            .SQL = "SELECT * FROM QWK_YPMF060"
        End If
        
        .Caption = "入金一覧表"
        If .PrintActiveReport(0) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End With

    Set objRpt1 = Nothing
    Set objArPrint = Nothing

    Screen.MousePointer = vbDefault

    Exit Sub

cmdExecute_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("実行クリックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Form_KeyDown_Err

    'リターンキーで次のコントロールへフォーカス移動
    If KeyCode = vbKeyReturn And Shift <> vbShiftMask Then
        KeyCode = 0
        Global_SendKeys Me, VK_TAB
        Exit Sub
    End If
    
    Exit Sub

Form_KeyDown_Err:

    Call MsgBox("フォームキーダウン時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")
 
End Sub

Private Sub Form_Load()

    Dim strBeforeDate As String

    On Error GoTo Form_Load_Err

    strBeforeDate = Global_Get_BeforeOdate(g_clsReg.LDatabase, g_strOdate)
    strBeforeDate = CStr(Format(DateAdd("d", 1, CDate(strBeforeDate)), "yyyy/mm/dd"))

    imnRdate_Year(0).Text = left$(strBeforeDate, 4)
    imnRdate_Month(0).Text = Mid$(strBeforeDate, 6, 2)
    imnRdate_Day(0).Text = right$(strBeforeDate, 2)
    imnRdate_Year(1).Text = left$(g_strOdate, 4)
    imnRdate_Month(1).Text = Mid$(g_strOdate, 6, 2)
    imnRdate_Day(1).Text = right$(g_strOdate, 2)
    
    Exit Sub
    
Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(imnRdate_Year(0).Text) = "" Then
        strErrMsg = "年を入力してください。"
        imnRdate_Year(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Month(0).Text) = "" Then
        strErrMsg = "月を入力してください。"
        imnRdate_Month(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Day(0).Text) = "" Then
        strErrMsg = "日を入力してください。"
        imnRdate_Day(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Year(1).Text) = "" Then
        strErrMsg = "年を入力してください。"
        imnRdate_Year(1).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Month(1).Text) = "" Then
        strErrMsg = "月を入力してください。"
        imnRdate_Month(1).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(imnRdate_Day(1).Text) = "" Then
        strErrMsg = "日を入力してください。"
        imnRdate_Day(1).SetFocus
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function

Private Sub imnRdate_Yeary_GotFocus(Index As Integer)

    imnRdate_Year(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Year_LostFocus(Index As Integer)

    imnRdate_Year(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRdate_Month_GotFocus(Index As Integer)

    imnRdate_Month(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Month_LostFocus(Index As Integer)

    imnRdate_Month(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnRdate_Day_GotFocus(Index As Integer)

    imnRdate_Day(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnRdate_Day_LostFocus(Index As Integer)

    imnRdate_Day(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Function MakeWork() As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim strBuff As String
    
    On Error GoTo MakeWork_Err
    
    MakeWork = False
    
    Screen.MousePointer = vbHourglass
    
    m_strRdateFrom = imnRdate_Year(0).Text & "/" & Format(imnRdate_Month(0).Text, "00") & "/" & Format(imnRdate_Day(0).Text, "00")
    m_strRdateTo = imnRdate_Year(1).Text & "/" & Format(imnRdate_Month(1).Text, "00") & "/" & Format(imnRdate_Day(1).Text, "00")
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF060"
    g_clsAdoAccess.Connection.Execute strSQL
    
    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF060"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    'データオープン
    strSQL = "SELECT * FROM DT060" & _
             " WHERE Rdate BETWEEN '" & m_strRdateFrom & "' AND '" & m_strRdateTo & "'" & _
             " ORDER BY Bcode,Odate,Rdate"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = True Then
        Screen.MousePointer = vbDefault
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
        Exit Function
    End If
    
    Do While Not adoRecordset1.EOF
        wkRecordset.AddNew
        wkRecordset.Fields("Bcode") = adoRecordset1.Fields("Bcode")
        wkRecordset.Fields("Odate") = adoRecordset1.Fields("Odate")
        wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, wkRecordset.Fields("Bcode"), wkRecordset.Fields("Odate"), strBuff)
        wkRecordset.Fields("Rdate") = adoRecordset1.Fields("Rdate")
        wkRecordset.Fields("Rdiv") = adoRecordset1.Fields("Rdiv")
        wkRecordset.Fields("R") = adoRecordset1.Fields("R")
        Select Case wkRecordset.Fields("R")
            Case PAYMENT_DIV_CASH:
                wkRecordset.Fields("Rname") = "現金"
            Case PAYMENT_DIV_CHECK:
                wkRecordset.Fields("Rname") = "小切手"
            Case PAYMENT_DIV_TRANSFER:
                wkRecordset.Fields("Rname") = "銀行振込"
            Case Else
                wkRecordset.Fields("Rname") = "現金"
        End Select
        
        wkRecordset.Fields("Ptotal") = adoRecordset1.Fields("Ptotal")
        wkRecordset.Fields("Ptotal2") = adoRecordset1.Fields("Ptotal2")
        '201107
        wkRecordset.Fields("Ptotal3") = adoRecordset1.Fields("Ptotal3")
                
        wkRecordset.Update
        adoRecordset1.MoveNext
    Loop
        
    adoRecordset1.Close
    wkRecordset.Requery     'バグ防止
    wkRecordset.Close
    
    Screen.MousePointer = vbDefault
    
    MakeWork = True
    
    Exit Function
    
MakeWork_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork_Err")
    
End Function


