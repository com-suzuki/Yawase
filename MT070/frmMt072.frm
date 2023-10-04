VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmMt072 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "買主番号変更"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmMt072.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdAkiban 
      Caption         =   "空番表示"
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
      Left            =   3120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "変更"
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
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "戻る"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1920
      _Version        =   262145
      _ExtentX        =   3387
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "変更前買主コード"
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
      LabelWidth      =   112
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
   Begin imText6Ctl.imText txtBcodeBefore 
      Height          =   360
      Left            =   2100
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   635
      Caption         =   "frmMt072.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt072.frx":007A
      Key             =   "frmMt072.frx":0098
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
      Caption         =   "frmMt072.frx":00CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt072.frx":013A
      Key             =   "frmMt072.frx":0158
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
      TabIndex        =   6
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmMt072.frx":019C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt072.frx":020A
      Key             =   "frmMt072.frx":0228
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
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   1920
      _Version        =   262145
      _ExtentX        =   3387
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "変更後買主コード"
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
      LabelWidth      =   112
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
   Begin imText6Ctl.imText txtBcodeAfter 
      Height          =   360
      Left            =   2100
      TabIndex        =   2
      Top             =   1020
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   635
      Caption         =   "frmMt072.frx":026C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMt072.frx":02DA
      Key             =   "frmMt072.frx":02F8
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
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "HGPｺﾞｼｯｸE"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmMt072"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAkiban_Click()

    frmMt072Search.Show vbModal
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdExecute_Click()

    If MsgBox("買主番号を変更しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "") = vbNo Then Exit Sub
    If DoValidationChecks() = False Then Exit Sub
    If DataUpdate() = True Then
        MsgBox "終了しました。", vbOKOnly + vbInformation, "情報"
        Unload Me
    End If
    
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

    Call FieldsClear
    txtBcodeBefore.Text = frmMt070.txtBcode
    
End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdCancel.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    txtBcodeBefore.SetFocus
    
End Sub

Private Sub txtBcodeBefore_GotFocus()
    
    txtBcodeBefore.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtBcodeBefore_LostFocus()
    
    txtBcodeBefore.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtBcodeAfter_GotFocus()
    
    txtBcodeAfter.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtBcodeAfter_LostFocus()
    
    txtBcodeAfter.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：画面クリア
'条　件　　：
'結　果　　：
'引　数　　：0：全画面 1:明細部
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／０９／１２
'更新履歴　：
'
Private Sub FieldsClear()

    On Error GoTo FieldsClear_Err
    
    txtBcodeBefore.Text = ""
    txtBcodeAfter.Text = ""
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

'目　的　　：更新処理
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／０９／１２
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.RecordSet
    Dim strBname As String
    
    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    frmMt070.m_clsAdoSQL.Connection.BeginTrans
    
    strBname = ""
    With adoRecordset1
        '得意先マスタ
        strSQL = "{call sp_MT070;2(" & txtBcodeBefore.Text & ")}"
        .Open strSQL, frmMt070.m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Fdiv")) Then
                If .Fields("Fdiv") = BUSINESS_DIV_BUYER Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    strBname = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
 
    'マスタにないデータは削除してしまう
    strSQL = "DELETE MT071" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE DT011" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE DT021" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE DT031" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE DT041" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE DT060" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE RT011" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE RT021" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE RT031" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE RT041" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "DELETE RT060" & _
             " WHERE Bcode = " & txtBcodeAfter.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    'マスタ
    strSQL = "UPDATE MT070 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
    
    strSQL = "UPDATE MT071 SET Bcode = " & txtBcodeAfter.Text & _
             " ,Bnum = " & txtBcodeAfter.Text & _
             " ,Fdate = '2099/12/31'" & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
    
    'データ
    strSQL = "UPDATE DT011 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "UPDATE DT021 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "UPDATE DT031 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "UPDATE DT041 SET Bcode = " & txtBcodeAfter.Text & _
             " ,Bname = '" & strBname & "'" & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "UPDATE DT060 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
    
    '累積
    strSQL = "UPDATE RT011 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "UPDATE RT021 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "UPDATE RT031 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "UPDATE RT041 SET Bcode = " & txtBcodeAfter.Text & _
             " ,Bname = '" & strBname & "'" & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
 
    strSQL = "UPDATE RT060 SET Bcode = " & txtBcodeAfter.Text & _
             " WHERE Bcode = " & txtBcodeBefore.Text
    frmMt070.m_clsAdoSQL.Connection.Execute strSQL
    
    frmMt070.m_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    Set adoRecordset1 = Nothing
    
    DataUpdate = True
    
    Exit Function

DataUpdate_Err:

    frmMt070.m_clsAdoSQL.Connection.RollbackTrans
    DataUpdate = False
    Screen.MousePointer = vbDefault
    Call MsgBox("データ更新エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_Err")

End Function

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／０９／１２
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.RecordSet
    
    On Error GoTo DoValidationChecks_Err

    If Trim(txtBcodeBefore.Text) = "" Then
        strErrMsg = "変更前買主コードを入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtBcodeAfter.Text) = "" Then
        strErrMsg = "変更後買主コードを入力してください。"
        GoTo ErrorTrap:
    End If
    
    With adoRecordset1
        strSQL = "{call sp_MT070;2(" & txtBcodeBefore.Text & ")}"
        .Open strSQL, frmMt070.m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If .EOF Then
            .Close
            strErrMsg = "変更前買主コードがマスタに存在しません。"
            GoTo ErrorTrap:
        End If
        .Close
        
        strSQL = "{call sp_MT070;2(" & txtBcodeAfter.Text & ")}"
        .Open strSQL, frmMt070.m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            .Close
            strErrMsg = "変更後買主コードが既に使われている番号を指定しました。"
            GoTo ErrorTrap:
        End If
        .Close
    End With
    
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

