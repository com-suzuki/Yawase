VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmRLogin 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "担当者入力"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5700
   StartUpPosition =   2  '画面の中央
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   60
      Top             =   1740
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1560
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4140
      TabIndex        =   6
      Top             =   1920
      Width           =   1395
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   1395
   End
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   5
      Left            =   60
      TabIndex        =   7
      Top             =   660
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
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
      _Version        =   262145
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "パスワード"
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
      LabelWidth      =   68
      LabelHeight     =   25
      LabelLeft       =   14
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
   Begin CSComboLib.CSComboBox cboPcode 
      Height          =   435
      Left            =   1560
      TabIndex        =   3
      Top             =   660
      Width           =   735
      _Version        =   262145
      _ExtentX        =   1296
      _ExtentY        =   767
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColDelim        =   ";"
      ColWidths       =   "2;20"
      Contents        =   "frmRLogin.frx":0000
      Extended        =   -1  'True
      ListBoxWidth    =   300
      MaxLength       =   2
      Text            =   "99"
   End
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   120
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
   Begin imText6Ctl.imText txtOdate_Year 
      Height          =   420
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   741
      Caption         =   "frmRLogin.frx":0019
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRLogin.frx":0087
      Key             =   "frmRLogin.frx":00A5
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
   Begin imText6Ctl.imText txtOdate_Month 
      Height          =   420
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   741
      Caption         =   "frmRLogin.frx":00D9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRLogin.frx":0147
      Key             =   "frmRLogin.frx":0165
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
   Begin CSComboLib.CSComboBox cboOdate_Day 
      Height          =   435
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   735
      _Version        =   262145
      _ExtentX        =   1296
      _ExtentY        =   767
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColDelim        =   ";"
      ColWidths       =   "2"
      Contents        =   "frmRLogin.frx":0199
      Extended        =   -1  'True
      MaxLength       =   2
      Text            =   "99"
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
      Left            =   2520
      TabIndex        =   13
      Top             =   180
      Width           =   375
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
      Left            =   3660
      TabIndex        =   12
      Top             =   180
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "日"
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
      Index           =   2
      Left            =   4980
      TabIndex        =   11
      Top             =   180
      Width           =   495
   End
   Begin VB.Label lblPname 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Caption         =   "ＮＮＮＮＮＮＮＮＮＮ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   2340
      TabIndex        =   9
      Top             =   660
      Width           =   3165
   End
End
Attribute VB_Name = "frmRLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboOdate_Day_GotFocus()

    cboOdate_Day.BackColor = FOCUS_STOP_COLOR
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboOdate_Day_LostFocus()

    cboOdate_Day.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboPcode_Click()

    Call cboPcode_Validate(False)

End Sub

Private Sub cboPcode_DropDown()

    Call Makecbo_Pcode(cboPcode)
    
End Sub

Private Sub cboPcode_GotFocus()

    cboPcode.BackColor = FOCUS_STOP_COLOR
    Call SetImeMode(ActiveControl.hwnd, 2)

End Sub

Private Sub cboPcode_LostFocus()

    cboPcode.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub cboPcode_Validate(Cancel As Boolean)

    If IsNumeric(cboPcode.Text) = False Then
        cboPcode.Text = ""
        lblPname.Caption = ""
        Exit Sub
    End If
    lblPname.Caption = Get_Pname(cboPcode.Text)
    If Trim(lblPname.Caption) = "" Then
        cboPcode.Text = ""
        cboPcode.Value = ""
    End If

End Sub

'目　的　　：
'条　件　　：キャンセルボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１２
'更新履歴　：
'
Private Sub cmdCancel_Click()

    Me.Hide

End Sub

'目　的　　：
'条　件　　：ＯＫボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１２
'更新履歴　：
'
Private Sub cmdOk_Click()

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strUser As String
    Dim strPassword As String

    On Error GoTo cmdOk_Click_Err
    
    If DoValidationChecks() = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    With adoRecordset1
        strSQL = "{call sp_MT030;2(" & cboPcode.Text & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then
            cboPcode.SetFocus
            Screen.MousePointer = vbDefault
            Call MsgBox("担当者が見つかりません。", vbOKOnly + vbCritical, "エラー")
            Exit Sub
        End If
        If IsNull(.Fields("PWD")) Then
            If Trim(txtPassword.Text) <> "" Then
                txtPassword.SetFocus
                Screen.MousePointer = vbDefault
                Call MsgBox("パスワードが違います。", vbOKOnly + vbCritical, "エラー")
                Exit Sub
            End If
        Else
            If Trim(.Fields("PWD")) <> Trim(txtPassword.Text) Then
                txtPassword.SetFocus
                Screen.MousePointer = vbDefault
                Call MsgBox("パスワードが違います。", vbOKOnly + vbCritical, "エラー")
                Exit Sub
            End If
        End If
    End With
    
    g_strPcode = Trim(cboPcode.Text)
    g_strPname = Trim(lblPname.Caption)
    g_strOdate = Trim(txtOdate_Year.Text) & "/" & Format(Trim(txtOdate_Month.Text), "00") & "/" & Format(Trim(cboOdate_Day.Text), "00")
    
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
    Screen.MousePointer = vbDefault
    
    g_blnLoginOK = True
    Me.Hide

    Exit Sub

cmdOk_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("ＯＫボタンクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdOk_Click_Err")

End Sub

'目　的　　：
'条　件　　：フォームロード時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１２
'更新履歴　：
'
Private Sub Form_Load()

    Dim intIndex1 As Integer

    On Error GoTo Form_Load_Err
    
    If Trim(g_strOdate) <> "" Then
        txtOdate_Year.Text = left$(g_strOdate, 4)
        txtOdate_Month.Text = Mid$(g_strOdate, 6, 2)
    Else
        txtOdate_Year.Text = Format(Now(), "yyyy")
        txtOdate_Month.Text = Format(Now(), "m")
    End If
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    txtPassword.Text = ""
    
    '開催日のコンボボックス作成
    cboOdate_Day.Clear
    For intIndex1 = 0 To UBound(HOLDING_DATE)
        cboOdate_Day.AddItem HOLDING_DATE(intIndex1)
    Next intIndex1
    
    Timer1.Enabled = True
    
    Exit Sub
    
Form_Load_Err:
    
    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    End
    
End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１２
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

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１２
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
    Dim intIndex1 As Integer
    Dim blnFlg As Boolean
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtOdate_Year.Text) = "" Then
        strErrMsg = "開催年月日を入力してください。"
        txtOdate_Year.SetFocus
        GoTo ErrorTrap:
    End If
    If Len(txtOdate_Year.Text) < 4 Then
        strErrMsg = "西暦年４桁で入力してください。"
        txtOdate_Year.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtOdate_Month.Text) = "" Then
        strErrMsg = "開催年月日を入力してください。"
        txtOdate_Month.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(cboOdate_Day.Text) = "" Then
        strErrMsg = "開催年月日を入力してください。"
        cboOdate_Day.SetFocus
        GoTo ErrorTrap:
    End If
    If IsDate(Trim(txtOdate_Year.Text) & "/" & Format(Trim(txtOdate_Month.Text), "00") & "/" & Format(Trim(cboOdate_Day.Text), "00")) = False Then
        strErrMsg = "正しい開催年月日を入力してください。"
        txtOdate_Year.SetFocus
        GoTo ErrorTrap:
    End If
    blnFlg = False
    For intIndex1 = 0 To cboOdate_Day.ListCount - 1
        If Trim(cboOdate_Day.List(intIndex1)) = Trim(cboOdate_Day.Text) Then
            blnFlg = True
            Exit For
        End If
    Next intIndex1
    If blnFlg = False Then
        strErrMsg = "正しい開催日を入力してください。"
        cboOdate_Day.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(cboPcode.Text) = "" Then
        strErrMsg = "担当者コードを入力してください。"
        cboPcode.SetFocus
        GoTo ErrorTrap:
    End If
    If IsNumeric(cboPcode.Text) = False Then
        strErrMsg = "正しい担当者コードを入力してください。"
        cboPcode.SetFocus
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

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Sub Makecbo_Pcode(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo Makecbo_Pcode_Err
    
    strBuff1 = Ctrl.Text
    Ctrl.Clear
    
    With adoRecordset1
        strSQL = "{call sp_MT030;1}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Ctrl.AddItem .Fields("Pcode") & ";" & .Fields("Pname")
            .MoveNext
        Loop
    End With
    
    Ctrl.Text = strBuff1
    
    Exit Sub
    
Makecbo_Pcode_Err:

    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Makecbo_Pcode_Err")

End Sub

'目　的　　：名称の取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／１０
'更新履歴　：
'
Private Function Get_Pname(strCode As String) As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo Get_Pname_Err
    
    Get_Pname = ""
    
    If Trim(strCode) = "" Then Exit Function
    
    With adoRecordset1
        strSQL = "{call sp_MT030;2(" & strCode & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            Get_Pname = IIf(IsNull(.Fields("Pname")), "", .Fields("Pname"))
        End If
    End With
    
    Exit Function
    
Get_Pname_Err:

    Call MsgBox("名称取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Get_Pname_Err")

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then Me.Hide

End Sub

Private Sub Timer1_Timer()
    
    Call Global_SetFormTop(Me.hwnd, True)
    Timer1.Enabled = False

End Sub

Private Sub txtOdate_Month_GotFocus()

    txtOdate_Month.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtOdate_Month_LostFocus()

    txtOdate_Month.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtOdate_Year_GotFocus()

    txtOdate_Year.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtOdate_Year_LostFocus()

    txtOdate_Year.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtPassword_GotFocus()

    txtPassword.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtPassword_LostFocus()

    txtPassword.BackColor = FOCUS_NO_COLOR
    
End Sub
