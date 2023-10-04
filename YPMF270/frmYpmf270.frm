VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmYpmf270 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   2205
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf270.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7110
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   6975
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'なし
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   720
         Width           =   2175
         Begin VB.OptionButton optDiv 
            Caption         =   "社外"
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
            TabIndex        =   4
            Top             =   120
            Width           =   915
         End
         Begin VB.OptionButton optDiv 
            Caption         =   "社内"
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
            TabIndex        =   3
            Top             =   120
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "集計年月"
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
      Begin imText6Ctl.imText txtYear 
         Height          =   420
         Left            =   1680
         TabIndex        =   1
         Top             =   300
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   741
         Caption         =   "frmYpmf270.frx":0CFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf270.frx":0D68
         Key             =   "frmYpmf270.frx":0D86
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
         Height          =   420
         Left            =   3120
         TabIndex        =   2
         Top             =   300
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   741
         Caption         =   "frmYpmf270.frx":0DBA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf270.frx":0E28
         Key             =   "frmYpmf270.frx":0E46
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
         Index           =   12
         Left            =   180
         TabIndex        =   15
         Top             =   840
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "区　分"
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
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3780
         TabIndex        =   13
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   9
      Top             =   1380
      Width           =   6975
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   5220
         TabIndex        =   6
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
         Picture         =   "frmYpmf270.frx":0E7A
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   3480
         TabIndex        =   5
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
         Picture         =   "frmYpmf270.frx":0FD4
      End
   End
   Begin imText6Ctl.imText imtFocusEnd 
      Height          =   135
      Left            =   7380
      TabIndex        =   7
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf270.frx":10E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf270.frx":1154
      Key             =   "frmYpmf270.frx":1172
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
      TabIndex        =   8
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf270.frx":11B6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf270.frx":1224
      Key             =   "frmYpmf270.frx":1242
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
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   7200
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf270.frx":1286
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf270.frx":12F4
      Key             =   "frmYpmf270.frx":1312
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
Attribute VB_Name = "frmYpmf270"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsAdoSQL As New clsAdoCore
Private m_clsAdoAccess As New clsAdoCore
Private m_clsReg As New clsReg

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２１
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

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
'作成年月日：２００２／０８／２１
'更新履歴　：
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２１
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
'作成年月日：２００２／０８／２１
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "植木流通動向調査表２"

    '重複起動のチェック
    If App.PrevInstance = True Then
        End
    End If

    'フォームクリア
    txtYear.Text = Format(Now(), "yyyy")
    txtMonth.Text = Format(Now(), "m")
    
    'レジストリ読み込み
    m_clsReg.RegKey = REG_KEY
    If m_clsReg.ReadReg = False Then
        End
    End If

    'データベース接続
    With m_clsAdoSQL
        .Provider = adoSQLServer
        .Server = m_clsReg.Server
        .DBName = m_clsReg.DBName
        .UID = m_clsReg.UID
        .PWD = m_clsReg.PWD
        .CommandTimeOut = m_clsReg.CommandTimeOut
        If .Connect = False Then
            End
        End If
    End With
    With m_clsAdoAccess
        .Provider = adoAccess
        .DBName = m_clsReg.LDatabase & "\" & m_clsReg.LDBName
        If .Connect = False Then
            End
        End If
    End With
    
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
'作成年月日：２００２／０８／２１
'更新履歴　：
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set m_clsAdoSQL = Nothing
    Set m_clsAdoAccess = Nothing
    Set m_clsReg = Nothing
    End
    
    Exit Sub
    
Form_Unload_Err:

    Call MsgBox("フォームアンロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")
    End

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cmdExecute.SetFocus

End Sub

Private Sub txtMonth_GotFocus()

    txtMonth.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtMonth_LostFocus()

    txtMonth.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtYear_GotFocus()

    txtYear.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtYear_LostFocus()

    txtYear.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：印刷用ワーク作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２１
'更新履歴　：
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim adoCmd As New ADODB.Command
    Dim adoParam As ADODB.Parameter
    Dim strYyyymm As String
    Dim intKaisu As Integer

    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    '集計年月
    strYyyymm = txtYear.Text & "/" & Format(txtMonth.Text, "00")
    
    '設定マスタオープン
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = True Then
        Call MsgBox("設定マスタが設定されていません。", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If

    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF270"
    m_clsAdoAccess.Connection.Execute strSQL

    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF270"
    wkRecordset.Open strSQL, m_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    'ワーク作成
    wkRecordset.AddNew
    wkRecordset.Fields("Yyyymm") = txtYear.Text & "年" & Format(txtMonth.Text, "00") & "月"
    wkRecordset.Fields("Company") = adoMT010.Fields("Company")
    wkRecordset.Fields("Tm2") = adoMT010.Fields("Tm2")
    wkRecordset.Fields("Om2") = adoMT010.Fields("Om2")
    wkRecordset.Fields("Im2") = adoMT010.Fields("Im2")
    wkRecordset.Fields("Cm2") = adoMT010.Fields("Cm2")
    wkRecordset.Fields("Bm2") = adoMT010.Fields("Bm2")
    
    '売上金額（買主精算データから取得）
    With adoCmd
        .ActiveConnection = m_clsAdoSQL.Connection
        '戻り値
        Set adoParam = .CreateParameter("rtn001", adCurrency, adParamReturnValue)
        .Parameters.Append adoParam
        '引数
        Set adoParam = .CreateParameter("arg001", adChar, adParamInput, 7)
        .Parameters.Append adoParam
    
        .CommandText = "sp_YPMF2701;1"
        .Parameters("arg001") = strYyyymm
        .CommandType = adCmdStoredProc
        .Execute

        wkRecordset.Fields("Uriage_Kingaku") = .Parameters("rtn001")
    End With
    Set adoParam = Nothing
    Set adoCmd = Nothing
    
    '売上本数
    With adoCmd
        .ActiveConnection = m_clsAdoSQL.Connection
        '引数
        Set adoParam = .CreateParameter("arg001", adChar, adParamInput, 6)
        .Parameters.Append adoParam
        Set adoParam = .CreateParameter("arg002", adChar, adParamInput, 7)
        .Parameters.Append adoParam
        Set adoParam = .CreateParameter("rtn001", adCurrency, adParamOutput)
        .Parameters.Append adoParam
        Set adoParam = .CreateParameter("rtn002", adCurrency, adParamOutput)
        .Parameters.Append adoParam
    
        .CommandText = "sp_YPMF2702;1"
        .Parameters("arg001") = txtYear.Text & Format(txtMonth.Text, "00")
        .Parameters("arg002") = txtYear.Text & "/" & Format(txtMonth.Text, "00")
        .CommandType = adCmdStoredProc
        .Execute

        wkRecordset.Fields("Ueki") = .Parameters("rtn001")
        wkRecordset.Fields("Hachimono") = .Parameters("rtn002")
    End With
    Set adoParam = Nothing
    Set adoCmd = Nothing
    
'    '開催回数（受付データから取得）
'    With adoCmd
'        .ActiveConnection = m_clsAdoSQL.Connection
'        '戻り値
'        Set adoParam = .CreateParameter("rtn001", adInteger, adParamReturnValue)
'        .Parameters.Append adoParam
'        '引数
'        Set adoParam = .CreateParameter("arg001", adChar, adParamInput, 7)
'        .Parameters.Append adoParam
'
'        .CommandText = "sp_YPMF2703;1"
'        .Parameters("arg001") = strYyyymm
'        .CommandType = adCmdStoredProc
'        .Execute
'
'        wkRecordset.Fields("Kaisu") = .Parameters("rtn001")
'    End With
'    Set adoParam = Nothing
'    Set adoCmd = Nothing
    
    '月別開催回数取得
    Call MONTH_HOLDING_DATE(CInt(txtMonth.Text), intKaisu)
    wkRecordset.Fields("Kaisu") = intKaisu
    
    wkRecordset.Update
    
    wkRecordset.Requery
    wkRecordset.Close
    
    MakePrintWork = True
    
MakePrintWork_Exit:
    
    Screen.MousePointer = vbDefault
    
    Exit Function

MakePrintWork_Cancel:

    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    MakePrintWork = False
    Call MsgBox("印刷ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function

'目　的　　：ActiveReportの印刷
'条　件　　：
'結　果　　：
'引　数　　：0:プレビュー 1:印刷
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２１
'更新履歴　：
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf270
    Dim objRpt_2 As New rptYpmf270_2
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "植木流通動向調査表２"
        If optDiv(0).Value = True Then
            .objReport = objRpt
        Else
            .objReport = objRpt_2
        End If
        .Connection = m_clsAdoAccess.Connection
        .Caption = "植木流通動向調査表２"
        If .PrintActiveReport(intFlg) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End With

    Set objRpt = Nothing
    Set objRpt_2 = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    ActiveReportPrint = True
    
    Exit Function
    
ActiveReportPrint_Err:

    ActiveReportPrint = False
    Screen.MousePointer = vbDefault
    Call MsgBox("ActiveReportの印刷エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReportPrint_Err")
    
End Function

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２１
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtYear.Text) = "" Then
        txtYear.SetFocus
        strErrMsg = "集計年月を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtMonth.Text) = "" Then
        txtMonth.SetFocus
        strErrMsg = "集計年月を入力してください。"
        GoTo ErrorTrap:
    End If
    If CInt(txtYear.Text) < 1900 Or CInt(txtYear.Text) > 2099 Then
        txtYear.SetFocus
        strErrMsg = "西暦年４桁で入力してください。"
        GoTo ErrorTrap:
    End If
    If CInt(txtMonth.Text) < 1 Or CInt(txtMonth.Text) > 12 Then
        txtMonth.SetFocus
        strErrMsg = "正しい集計年月を入力してください。"
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


