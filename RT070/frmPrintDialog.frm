VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmPrintDialog 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "印刷"
   ClientHeight    =   2610
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10335
   Icon            =   "frmPrintDialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   10335
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
      Height          =   1695
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   8835
      Begin VB.OptionButton optOrder 
         Caption         =   "はがき宛名(くじ付はがき)"
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
         Index           =   2
         Left            =   4560
         TabIndex        =   16
         Top             =   1200
         Width           =   3435
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "はがき宛名"
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
         Left            =   2880
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "一覧表"
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
         TabIndex        =   12
         Top             =   1200
         Value           =   -1  'True
         Width           =   1155
      End
      Begin imText6Ctl.imText txtBcode 
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   635
         Caption         =   "frmPrintDialog.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":007A
         Key             =   "frmPrintDialog.frx":0098
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
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "得意先コード"
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
         LabelWidth      =   82
         LabelHeight     =   25
         LabelLeft       =   7
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
         Index           =   1
         Left            =   2820
         TabIndex        =   5
         Top             =   360
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   635
         Caption         =   "frmPrintDialog.frx":00CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":013A
         Key             =   "frmPrintDialog.frx":0158
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
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   780
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "業務区分"
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'なし
         Height          =   435
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   4335
         Begin VB.OptionButton optFdiv 
            Caption         =   "両　方"
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
            Index           =   2
            Left            =   2880
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   11
            Top             =   60
            Width           =   1395
         End
         Begin VB.OptionButton optFdiv 
            Caption         =   "買主"
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
            Index           =   0
            Left            =   0
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   10
            Top             =   60
            Width           =   1395
         End
         Begin VB.OptionButton optFdiv 
            Caption         =   "出品者"
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
            Index           =   1
            Left            =   1440
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   9
            Top             =   60
            Width           =   1395
         End
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "業務区分"
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
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "登録日"
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
      Begin imDate6Ctl.imDate imdAdddate 
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   19
         Top             =   360
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         Calendar        =   "frmPrintDialog.frx":018C
         Caption         =   "frmPrintDialog.frx":030C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":037A
         Keys            =   "frmPrintDialog.frx":0398
         Spin            =   "frmPrintDialog.frx":03F6
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "yyyy/mm/dd"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "yyyy/mm/dd"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "2005/09/03"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   38598
         CenturyMode     =   0
      End
      Begin imDate6Ctl.imDate imdAdddate 
         Height          =   375
         Index           =   1
         Left            =   6960
         TabIndex        =   20
         Top             =   360
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         Calendar        =   "frmPrintDialog.frx":041E
         Caption         =   "frmPrintDialog.frx":059E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":060C
         Keys            =   "frmPrintDialog.frx":062A
         Spin            =   "frmPrintDialog.frx":0688
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "yyyy/mm/dd"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "yyyy/mm/dd"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "2005/09/03"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   38598
         CenturyMode     =   0
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
         Index           =   1
         Left            =   6660
         TabIndex        =   21
         Top             =   420
         Width           =   375
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
         Left            =   2460
         TabIndex        =   6
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
      Left            =   9000
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
      Left            =   9000
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "※年賀はがきなどは「くじ付はがき」を選択してください。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   60
      TabIndex        =   17
      Top             =   2220
      Width           =   9075
   End
   Begin VB.Label Label2 
      Caption         =   "※ハガキ宛名は、「上様」、郵便番号、住所、名前が空白の買主は印刷されません。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   60
      TabIndex        =   15
      Top             =   1860
      Width           =   9075
   End
End
Attribute VB_Name = "frmPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objArPrint As New clsArPrint

Private Sub cmdExecute_Click()
    
    Dim objRpt1 As New rptMT070
    Dim objRpt2 As New rptMT070_2
    Dim objRpt3 As New rptMT070_3
    Dim strSQL As String
        
    On Error GoTo cmdExecute_Click_Err
    
    If DoValidationChecks() = False Then Exit Sub
    
    cmdCancel.SetFocus
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        If optOrder(0).Value = True Then
            .Name = "得意先マスタ一覧表(過去分参照)"
            .objReport = objRpt1
            .Connection = frmMt070.m_clsAdoSQL.Connection
            If optFdiv(0).Value = True Then
                strSQL = "SELECT * FROM vw_MT070" & _
                         " WHERE Bcode BETWEEN '" & txtBcode(0).Text & "' AND '" & txtBcode(1).Text & "'" & _
                         " AND (Fdiv = " & BUSINESS_DIV_BUYER & " OR Fdiv = " & BUSINESS_DIV_ALL & ")"
            ElseIf optFdiv(1).Value = True Then
                strSQL = "SELECT * FROM vw_MT070" & _
                         " WHERE Bcode BETWEEN '" & txtBcode(0).Text & "' AND '" & txtBcode(1).Text & "'" & _
                         " AND (Fdiv = " & BUSINESS_DIV_EXHIBITION & " OR Fdiv = " & BUSINESS_DIV_ALL & ")"
            ElseIf optFdiv(2).Value = True Then
                strSQL = "SELECT * FROM vw_MT070" & _
                         " WHERE Bcode BETWEEN '" & txtBcode(0).Text & "' AND '" & txtBcode(1).Text & "'"
            End If
            If imdAdddate(0).ValueIsNull = False Then
                strSQL = strSQL & " AND CONVERT(VARCHAR(10),Adddate,111) >= '" & imdAdddate(0).Text & "'"
            End If
            If imdAdddate(1).ValueIsNull = False Then
                strSQL = strSQL & " AND CONVERT(VARCHAR(10),Adddate,111) <= '" & imdAdddate(1).Text & "'"
            End If
            .SQL = strSQL
            .Caption = "得意先マスタ一覧表(過去分参照)"
            If .PrintActiveReport(0) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            .Name = "ハガキ宛名"
            If optOrder(1).Value = True Then
                .objReport = objRpt2
            ElseIf optOrder(2).Value = True Then
                .objReport = objRpt3
            End If
            .Connection = frmMt070.m_clsAdoSQL.Connection
            If optFdiv(0).Value = True Then
                strSQL = "SELECT * FROM vw_MT070" & _
                       " WHERE Bcode BETWEEN '" & txtBcode(0).Text & "' AND '" & txtBcode(1).Text & "'" & _
                       " AND (Fdiv = " & BUSINESS_DIV_BUYER & " OR Fdiv = " & BUSINESS_DIV_ALL & ")" & _
                       " AND Bname IS NOT NULL AND Bname <> '上様' AND Post IS NOT NULL AND Addres IS NOT NULL"
            ElseIf optFdiv(1).Value = True Then
                strSQL = "SELECT * FROM vw_MT070" & _
                       " WHERE Bcode BETWEEN '" & txtBcode(0).Text & "' AND '" & txtBcode(1).Text & "'" & _
                       " AND (Fdiv = " & BUSINESS_DIV_EXHIBITION & " OR Fdiv = " & BUSINESS_DIV_ALL & ")" & _
                       " AND Bname IS NOT NULL AND Bname <> '上様' AND Post IS NOT NULL AND Addres IS NOT NULL"
            ElseIf optFdiv(2).Value = True Then
                strSQL = "SELECT * FROM vw_MT070" & _
                       " WHERE Bcode BETWEEN '" & txtBcode(0).Text & "' AND '" & txtBcode(1).Text & "'" & _
                       " AND Bname IS NOT NULL AND Bname <> '上様' AND Post IS NOT NULL AND Addres IS NOT NULL"
            End If
            If imdAdddate(0).ValueIsNull = False Then
                strSQL = strSQL & " AND CONVERT(VARCHAR(10),Adddate,111) >= '" & imdAdddate(0).Text & "'"
            End If
            If imdAdddate(1).ValueIsNull = False Then
                strSQL = strSQL & " AND CONVERT(VARCHAR(10),Adddate,111) <= '" & imdAdddate(1).Text & "'"
            End If
            .SQL = strSQL
            .Caption = "ハガキ宛名"
            If .PrintActiveReport(0) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End With
    
    Set objRpt1 = Nothing
    Set objRpt2 = Nothing
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

    txtBcode(0).Text = "0"
    txtBcode(1).Text = "9999"
    optFdiv(0).Value = True
    imdAdddate(0).Value = Null
    imdAdddate(1).Value = Null
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtBcode(0).Text) = "" Then
        strErrMsg = "得意先コードを入力してください。"
        txtBcode(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtBcode(1).Text) = "" Then
        strErrMsg = "得意先コードを入力してください。"
        txtBcode(1).SetFocus
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

Private Sub imdAdddate_GotFocus(Index As Integer)

    imdAdddate(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imdAdddate_LostFocus(Index As Integer)

    imdAdddate(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtBcode_GotFocus(Index As Integer)

    txtBcode(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtBcode_LostFocus(Index As Integer)

    txtBcode(Index).BackColor = FOCUS_NO_COLOR
    
End Sub
