VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmYpmf140 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   3105
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
   Icon            =   "frmYpmf140.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   12150
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdLogin 
         Caption         =   "開催年月日と担当者の変更"
         Height          =   375
         Left            =   6960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9900
         TabIndex        =   11
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
         Contents        =   "frmYpmf140.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   9
      Top             =   2280
      Width           =   12015
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   3
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
         Picture         =   "frmYpmf140.frx":0D13
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10200
         TabIndex        =   5
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
         Picture         =   "frmYpmf140.frx":0D2F
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   8460
         TabIndex        =   4
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
         Picture         =   "frmYpmf140.frx":0E89
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   1635
      Left            =   60
      TabIndex        =   8
      Top             =   660
      Width           =   12015
      Begin CSComboLib.CSComboBox cboBcode 
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
         Contents        =   "frmYpmf140.frx":0F9B
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   180
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
      Begin CSComboLib.CSComboBox cboBcode 
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
         Contents        =   "frmYpmf140.frx":0FB4
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
         TabIndex        =   20
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblBcode_Name 
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
         TabIndex        =   19
         Top             =   1020
         Width           =   9195
      End
      Begin VB.Label lblBcode_Name 
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
         TabIndex        =   18
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
      Caption         =   "frmYpmf140.frx":0FCD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf140.frx":103B
      Key             =   "frmYpmf140.frx":1059
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
      TabIndex        =   6
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf140.frx":109D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf140.frx":110B
      Key             =   "frmYpmf140.frx":1129
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
      TabIndex        =   7
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf140.frx":116D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf140.frx":11DB
      Key             =   "frmYpmf140.frx":11F9
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
Attribute VB_Name = "frmYpmf140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_strGinkoName  As String

Private Sub cboBcode_Click(Index As Integer)

    Call cboBcode_Validate(Index, False)
    
End Sub

Private Sub cboBcode_DropDown(Index As Integer)

    Call MakecboBcode(cboBcode(Index))
    
End Sub

Private Sub cboBcode_GotFocus(Index As Integer)

    cboBcode(Index).BackColor = FOCUS_STOP_COLOR
    cboBcode(Index).Tag = cboBcode(Index).Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboBcode_LostFocus(Index As Integer)
   
    cboBcode(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboBcode_Validate(Index As Integer, Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cboBcode_Validate_Err
    
    If Trim(cboBcode(Index).Text) = "" Then Exit Sub
    If cboBcode(Index).Tag = cboBcode(Index).Text Then Exit Sub
    
    lblBcode_Name(Index).Caption = ""
    
    With adoRecordset1
        '得意先マスタ
        strSQL = "{call sp_MT070;2(" & Trim(cboBcode(Index).Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Fdiv")) Then
                If .Fields("Fdiv") = BUSINESS_DIV_BUYER Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    lblBcode_Name(Index).Caption = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If Index = 0 Then
        cboBcode(1).Text = cboBcode(0).Text
        lblBcode_Name(1).Caption = lblBcode_Name(0).Caption
    End If
    
    Exit Sub

cboBcode_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBcode_Validate_Err")

End Sub

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００３／０６／１８
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    cboBcode(0).SetFocus

End Sub

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００３／０６／１８
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
'作成年月日：２００３／０６／１８
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
'作成年月日：２００３／０６／１８
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
'作成年月日：２００３／０６／１８
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "未収分請求書出力"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    Call FieldsClear(0)
    
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
'作成年月日：２００３／０６／１８
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
'作成年月日：２００３／０６／１８
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        cboBcode(0).Text = "0"
        cboBcode(0).Tag = ""
        cboBcode(1).Text = "9999"
        cboBcode(1).Tag = ""
        lblBcode_Name(0).Caption = ""
        lblBcode_Name(1).Caption = ""
'        Call FieldsSet
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
'作成年月日：２００３／０６／１８
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "開催年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(cboBcode(0).Text) = "" Then
        cboBcode(0).SetFocus
        strErrMsg = "買主コードを入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(cboBcode(1).Text) = "" Then
        cboBcode(1).SetFocus
        strErrMsg = "買主コードを入力してください。"
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
'作成年月日：２００３／０６／１８
'更新履歴　：
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim adoDT041 As New ADODB.Recordset
    Dim adoDT041_M As New ADODB.Recordset
    Dim adoDT060 As New ADODB.Recordset
    Dim adoDT021 As New ADODB.Recordset
    Dim adoMT010 As New ADODB.Recordset
    Dim adoMT070 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim strBuff As String
    Dim intIndex1 As Integer

    Dim intPage As Integer
    Dim intLine As Integer

    Dim strKey As String
    Dim strBcode As String
    Dim strBname As String
    Dim strPost As String
    Dim strAdress As String
    Dim curSeikyu_Total As Currency
    Dim curPrice_Total As Currency
    Dim curNyukin_Total As Currency
    Dim strKessai_Date As String
    
    Const PAGE_MAX_LINE = 24                    '1ページの最大行数

    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    '会社設定マスタ
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        m_strGinkoName = RTrim(adoMT010("Memo"))
    Else
        m_strGinkoName = ""
    End If
    adoMT010.Close
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF140"
    g_clsAdoAccess.Connection.Execute strSQL

    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF140"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '買主精算データオープン
    strSQL = "{call sp_YPMF1401;1('" & Trim(lblOdate.Caption) & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    adoDT041.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT041.EOF = True Then
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If
        
    frmCount.fpProgressBar1.Value = 0
    frmCount.fpProgressBar1.Max = adoDT041.RecordCount
    frmCount.Show
    Me.Enabled = False
    
    Do While Not adoDT041.EOF
        strBcode = Format$(adoDT041.Fields("Bcode"), "0000")
        strBname = Global_Get_Bname(g_clsAdoSQL, strBcode, Trim(lblOdate.Caption), strBuff) & "　様"
        
        '得意先マスタ
        strSQL = "{call sp_MT070;2(" & strBuff & ")}"
        adoMT070.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoMT070.EOF = False Then
            strPost = IIf(IsNull(adoMT070.Fields("Post")), "", Trim(adoMT070.Fields("Post")))
            strAdress = IIf(IsNull(adoMT070.Fields("Addres")), "", Trim(adoMT070.Fields("Addres")))
        Else
            strPost = ""
            strAdress = ""
        End If
        adoMT070.Close
        
        '初期化
        curSeikyu_Total = 0
        '売立総合計
        curPrice_Total = IIf(IsNull(adoDT041.Fields("Gtotal")), 0, adoDT041.Fields("Gtotal"))
        curNyukin_Total = 0
        strKessai_Date = Format$(Now(), "yyyy年mm月dd日") & "決済分"
        
        intPage = 1
        intLine = 1
        
        '買主精算データオープン
        strSQL = "{call sp_YPMF1402;1('" & Trim(lblOdate.Caption) & "'," & adoDT041.Fields("Bcode") & ")}"
        adoDT041_M.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT041_M.EOF
            
            'ヘッダー
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "開催日：" & Format$(adoDT041_M.Fields("Odate"), "yyyy年mm月dd日")
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = Null
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            'ページ計算
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
'********** 競売分 **********
            
            '競売明細データオープン
            strSQL = "{call sp_YPMF1403;1('" & Format$(adoDT041_M.Fields("Odate"), "yyyymmdd") & "'," & adoDT041_M.Fields("Bcode") & "," & adoDT041_M.Fields("Num") & ")}"
            adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT021.EOF Then
                adoDT021.Close
                
                'データがない場合は競売明細データ累積を探す
                strSQL = "{call sp_YPMF1403;2('" & Format$(adoDT041_M.Fields("Odate"), "yyyymmdd") & "'," & adoDT041_M.Fields("Bcode") & "," & adoDT041_M.Fields("Num") & ")}"
                adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            End If
            
            Do While Not adoDT021.EOF
                wkRecordset.AddNew
                wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
                wkRecordset.Fields("PageNum") = intPage
                wkRecordset.Fields("Line") = intLine
                
                wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
                wkRecordset.Fields("Bname") = strBname
                wkRecordset.Fields("Post") = strPost
                wkRecordset.Fields("Addres") = strAdress
                wkRecordset.Fields("Price_Total") = 0
                wkRecordset.Fields("Nyukin_Total") = 0
                wkRecordset.Fields("Kessai_Date") = strKessai_Date
                
                wkRecordset.Fields("Iname") = adoDT021.Fields("Iname")
                wkRecordset.Fields("Qty") = adoDT021.Fields("Qty")
                wkRecordset.Fields("Price") = adoDT021.Fields("Price")
                If adoDT021.Fields("Qty") <> 0 Then
                    wkRecordset.Fields("Tanka") = Fix(CCur(wkRecordset.Fields("Price")) / CCur(wkRecordset.Fields("Qty")))
                Else
                    wkRecordset.Fields("Tanka") = 0
                End If
                wkRecordset.Fields("Tekiyo") = ""
                wkRecordset.Update
                
                'ページ計算
                intLine = intLine + 1
                If intLine > PAGE_MAX_LINE Then
                    intPage = intPage + 1
                    intLine = 1
                End If
                
                adoDT021.MoveNext
            Loop
            adoDT021.Close
                    
'********** 注文分 **********
            
             '注文明細データ
            strSQL = "SELECT * FROM DT031" & _
                     " WHERE Odate = '" & adoDT041_M.Fields("Odate") & "'" & _
                     " AND Bcode = " & adoDT041_M.Fields("Bcode") & _
                     " AND Bnum = " & adoDT041_M.Fields("Num") & _
                     " ORDER BY Onum, Line"
            adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT021.EOF Then
                adoDT021.Close
            
                'データがない場合は注文明細データ累積を探す
                strSQL = "SELECT * FROM RT031" & _
                         " WHERE Odate = '" & adoDT041_M.Fields("Odate") & "'" & _
                         " AND Bcode = " & adoDT041_M.Fields("Bcode") & _
                         " AND Bnum = " & adoDT041_M.Fields("Num") & _
                         " ORDER BY Onum, Line"
                adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            End If
            
            Do While Not adoDT021.EOF
                wkRecordset.AddNew
                wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
                wkRecordset.Fields("PageNum") = intPage
                wkRecordset.Fields("Line") = intLine
                
                wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
                wkRecordset.Fields("Bname") = strBname
                wkRecordset.Fields("Post") = strPost
                wkRecordset.Fields("Addres") = strAdress
                wkRecordset.Fields("Price_Total") = 0
                wkRecordset.Fields("Nyukin_Total") = 0
                wkRecordset.Fields("Kessai_Date") = strKessai_Date
                
                wkRecordset.Fields("Iname") = adoDT021.Fields("Iname")
                wkRecordset.Fields("Qty") = adoDT021.Fields("Qty")
                wkRecordset.Fields("Price") = adoDT021.Fields("Price")
                If adoDT021.Fields("Qty") <> 0 Then
                    wkRecordset.Fields("Tanka") = Fix(CCur(wkRecordset.Fields("Price")) / CCur(wkRecordset.Fields("Qty")))
                Else
                    wkRecordset.Fields("Tanka") = 0
                End If
                wkRecordset.Fields("Tekiyo") = ""
                wkRecordset.Update
                
                'ページ計算
                intLine = intLine + 1
                If intLine > PAGE_MAX_LINE Then
                    intPage = intPage + 1
                    intLine = 1
                End If
                
                adoDT021.MoveNext
            Loop
            adoDT021.Close
                    
            '小計
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "           小　計"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Total")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            'ページ計算
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
            '消費税
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "           消費税"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Tax")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            'ページ計算
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
            '維持管理費
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "           維持管理費"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Keep")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            'ページ計算
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
            '201107 競売手数料
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "           競売手数料"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Brate2")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            'ページ計算
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
            '合計
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "********** 合　計 **********"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Gtotal")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            'ページ計算
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
        
            adoDT041_M.MoveNext
        Loop
        adoDT041_M.Close
        
        '買主精算データオープン
        strSQL = "SELECT Odate FROM DT041" & _
                 " WHERE Odate <= '" & Trim(lblOdate.Caption) & "'" & _
                 " AND Bcode = " & adoDT041.Fields("Bcode") & _
                 " AND Rdiv = 0 " & _
                 " GROUP BY Odate" & _
                 " ORDER BY Odate"
        adoDT041_M.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT041_M.EOF
            '入金データオープン
            strSQL = "{call sp_YPMF1404;1('" & adoDT041_M.Fields("Odate") & "'," & adoDT041.Fields("Bcode") & ")}"
            adoDT060.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            Do While Not adoDT060.EOF
                '入金
                wkRecordset.AddNew
                wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
                wkRecordset.Fields("PageNum") = intPage
                wkRecordset.Fields("Line") = intLine
                wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
                wkRecordset.Fields("Bname") = strBname
                wkRecordset.Fields("Post") = strPost
                wkRecordset.Fields("Addres") = strAdress
                wkRecordset.Fields("Price_Total") = 0
                wkRecordset.Fields("Nyukin_Total") = 0
                wkRecordset.Fields("Kessai_Date") = strKessai_Date
                
                Select Case adoDT060.Fields("R")
                    Case "1":
                        wkRecordset.Fields("Iname") = Format(adoDT060.Fields("Rdate"), "yyyy年mm月dd日") & "　入金(現金)"
                    Case "2":
                        wkRecordset.Fields("Iname") = Format(adoDT060.Fields("Rdate"), "yyyy年mm月dd日") & "　入金(小切手)"
                    Case "3":
                        wkRecordset.Fields("Iname") = Format(adoDT060.Fields("Rdate"), "yyyy年mm月dd日") & "　入金(銀行振込)"
                End Select
                
                wkRecordset.Fields("Qty") = Null
                wkRecordset.Fields("Price") = adoDT060.Fields("Ptotal")
                wkRecordset.Fields("Tanka") = Null
                wkRecordset.Fields("Tekiyo") = ""
                wkRecordset.Update
                
                'ページ計算
                intLine = intLine + 1
                If intLine > PAGE_MAX_LINE Then
                    intPage = intPage + 1
                    intLine = 1
                End If
            
                '入金金額計算
                curNyukin_Total = curNyukin_Total + CCur(adoDT060.Fields("Ptotal"))
                    
                adoDT060.MoveNext
            Loop
            adoDT060.Close
            
            adoDT041_M.MoveNext
        Loop
        adoDT041_M.Close
            
        '空行作成(最終行までデータがある場合はintLineが１となる)
        If intLine <> 1 Then
            For intIndex1 = intLine To PAGE_MAX_LINE
                wkRecordset.AddNew
                wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
                wkRecordset.Fields("PageNum") = intPage
                wkRecordset.Fields("Line") = intIndex1
                wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
                wkRecordset.Fields("Bname") = strBname
                wkRecordset.Fields("Post") = strPost
                wkRecordset.Fields("Addres") = strAdress
                wkRecordset.Fields("Price_Total") = 0
                wkRecordset.Fields("Nyukin_Total") = 0
                wkRecordset.Fields("Kessai_Date") = strKessai_Date
                wkRecordset.Fields("Iname") = Null
                wkRecordset.Fields("Qty") = Null
                wkRecordset.Fields("Price") = Null
                wkRecordset.Fields("Tanka") = Null
                wkRecordset.Fields("Tekiyo") = Null
                wkRecordset.Update
            Next intIndex1
        End If
       
        '請求金額を計算（売立総合計－入金合計）
        curSeikyu_Total = curPrice_Total - curNyukin_Total

        If curSeikyu_Total > 0 Then
            '請求額更新
            strSQL = "UPDATE WK_YPMF140"
            strSQL = strSQL & " SET Seikyu_Total = " & curSeikyu_Total & ","
            strSQL = strSQL & " Price_Total = " & curPrice_Total & ","
            strSQL = strSQL & " Nyukin_Total = " & curNyukin_Total
            strSQL = strSQL & " WHERE Bcode = '" & "(" & strBcode & ")" & "'"
            g_clsAdoAccess.Connection.Execute strSQL
        Else
            '残金がない場合はワークデータ削除
            strSQL = "DELETE FROM WK_YPMF140"
            strSQL = strSQL & " WHERE Bcode = '" & "(" & strBcode & ")" & "'"
            g_clsAdoAccess.Connection.Execute strSQL
        End If

        adoDT041.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    adoDT041.Close
    
    wkRecordset.Close
    
    'バグ防止
    strSQL = "SELECT * FROM WK_YPMF140"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    wkRecordset.Requery
    
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

    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    MakePrintWork = False
    Call MsgBox("印刷ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００３／０６／１８
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

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cboBcode(0).SetFocus

End Sub

'目　的　　：ActiveReportの印刷
'条　件　　：
'結　果　　：
'引　数　　：0:プレビュー 1:印刷
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００３／０６／１８
'更新履歴　：
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf140
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "未収分請求書"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .Caption = "未収分請求書"
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

'目　的　　：フィールドのセット
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００３／０６／１８
'更新履歴　：
'
Private Function FieldsSet() As Boolean
    
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim strBuff As String

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    '買主精算データ
    strSQL = "SELECT * FROM MT070" & _
             " WHERE Fdiv = " & BUSINESS_DIV_BUYER & " OR Fdiv = " & BUSINESS_DIV_ALL & _
             " ORDER BY Bcode"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        cboBcode(0).Text = adoRecordset1.Fields("Bcode")
        lblBcode_Name(0).Caption = Global_Get_Bname(g_clsAdoSQL, adoRecordset1.Fields("Bcode"), lblOdate.Caption, strBuff)
        adoRecordset1.MoveLast
        cboBcode(1).Text = adoRecordset1.Fields("Bcode")
        lblBcode_Name(1).Caption = Global_Get_Bname(g_clsAdoSQL, adoRecordset1.Fields("Bcode"), lblOdate.Caption, strBuff)
    End If
    
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
    Screen.MousePointer = vbDefault
    
    FieldsSet = True
    
    Exit Function

FieldsSet_Err:

    FieldsSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("フィールドセットエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsSet_Err")

End Function

