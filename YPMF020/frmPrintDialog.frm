VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmPrintDialog 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "印刷"
   ClientHeight    =   1470
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7170
   Icon            =   "frmPrintDialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   60
      Top             =   1440
   End
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
      Width           =   5655
      Begin VB.CheckBox chkReprint 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin imText6Ctl.imText txtOcode 
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   12
         LengthAsByte    =   0
         Text            =   "999999999999"
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
      Begin imText6Ctl.imText txtOcode 
         Height          =   360
         Index           =   1
         Left            =   3780
         TabIndex        =   5
         Top             =   360
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   635
         Caption         =   "frmPrintDialog.frx":00DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog.frx":014A
         Key             =   "frmPrintDialog.frx":0168
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   12
         LengthAsByte    =   0
         Text            =   "999999999999"
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
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   840
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "再発行"
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
         TabIndex        =   7
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
      Left            =   5820
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
      Left            =   5820
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
Public m_blnAutoPrint As Boolean

Private Sub cmdExecute_Click()
    
    Dim objRpt As New rptYpmf020
    
    On Error GoTo cmdExecute_Click_Err
    
    If DoValidationChecks() = False Then Exit Sub
    If MakeWork() = False Then Exit Sub
    
    If m_blnAutoPrint = False Then cmdCancel.SetFocus
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "競売結果確認表"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_YPMF020"
        .Caption = "競売結果確認表"
        If m_blnAutoPrint = True Then
            If .PrintActiveReport(1) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            If .PrintActiveReport(0) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End With
    
    Set objRpt = Nothing
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

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo Form_Load_Err

    txtOcode(0).Text = ""
    txtOcode(1).Text = ""
    chkReprint.Value = 0
    
    If m_blnAutoPrint = True Then
        chkReprint.Value = 1
        txtOcode(0).Text = frmYpmf020.m_strLastOcode
        txtOcode(1).Text = frmYpmf020.m_strLastOcode
        Me.Enabled = False
        Timer1.Enabled = True
        Exit Sub
    Else
        txtOcode(0).Text = "000000000000"
        With adoRecordset1
            strSQL = "SELECT * FROM DT020" & _
                 " WHERE Odate = '" & frmYpmf020.lblOdate.Caption & "'" & _
                 " ORDER BY Ocode DESC"
            .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If Not .EOF Then
                txtOcode(1).Text = .Fields("Ocode")
            Else
                txtOcode(1).Text = "999999999999"
            End If
        End With
    End If
    
    Exit Sub
    
Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtOcode(0).Text) = "" Then
        strErrMsg = "競売番号を入力してください。"
        txtOcode(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtOcode(1).Text) = "" Then
        strErrMsg = "競売番号を入力してください。"
        txtOcode(1).SetFocus
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

Private Sub Timer1_Timer()

    Timer1.Enabled = False
    Call cmdExecute_Click
    Unload Me

End Sub

Private Sub txtOcode_GotFocus(Index As Integer)

    txtOcode(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtOcode_LostFocus(Index As Integer)

    txtOcode(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Function MakeWork() As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim adoRecordset3 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim intLineCount As Integer
    Dim intIndex1 As Integer
    
    Const PAGE_MAX_ROW = 20
    
    On Error GoTo MakeWork_Err
    
    MakeWork = False
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF020"
    g_clsAdoAccess.Connection.Execute strSQL
    
    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF020"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    'データオープン
    strSQL = "SELECT * FROM DT020" & _
             " WHERE Odate = '" & frmYpmf020.lblOdate.Caption & "'" & _
             " AND Ocode BETWEEN '" & txtOcode(0).Text & "' AND '" & txtOcode(1).Text & "'"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        intLineCount = 1
        
        If chkReprint.Value = 0 Then
            'データオープン
            strSQL = "SELECT * FROM vw_YPMF020" & _
                     " WHERE Ocode = '" & adoRecordset1.Fields("Ocode") & "'" & _
                     " AND Wdiv = 0"
        Else
            'データオープン
            strSQL = "SELECT * FROM vw_YPMF020" & _
                     " WHERE Ocode = '" & adoRecordset1.Fields("Ocode") & "'"
        End If
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        Do While Not adoRecordset2.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Ocode") = adoRecordset2.Fields("Ocode")
            wkRecordset.Fields("Odate") = adoRecordset1.Fields("Odate")
            wkRecordset.Fields("Hnum") = adoRecordset2.Fields("Hnum")
            wkRecordset.Fields("Line") = intLineCount
            wkRecordset.Fields("Pnum") = adoRecordset2.Fields("Pnum")
            wkRecordset.Fields("PnumLine") = adoRecordset2.Fields("PnumLine")
            wkRecordset.Fields("Icode") = adoRecordset2.Fields("Icode")
            wkRecordset.Fields("Iname") = adoRecordset2.Fields("Iname")
            wkRecordset.Fields("Qty") = adoRecordset2.Fields("Qty")
            wkRecordset.Fields("Price") = adoRecordset2.Fields("Price")
            wkRecordset.Fields("Bcode") = adoRecordset2.Fields("Bcode")
            '備考欄
            If Not IsNull(adoRecordset2.Fields("Sline")) Then
                If CInt(adoRecordset2.Fields("Sline")) <> 0 Then
                    wkRecordset.Fields("Msg") = "合"
                End If
            Else
                wkRecordset.Fields("Msg") = ""
            End If
            If Not IsNull(adoRecordset2.Fields("Idiv")) Then
                If adoRecordset2.Fields("Idiv") = AUCTION_OFF Then
                    wkRecordset.Fields("Msg") = "ﾔﾒ"
                End If
            Else
                wkRecordset.Fields("Msg") = ""
            End If
            wkRecordset.Fields("Pcode") = adoRecordset2.Fields("Pcode")
            
            'データオープン
            strSQL = "SELECT Pname FROM MT030" & _
                     " WHERE Pcode = " & adoRecordset2.Fields("Pcode")
            adoRecordset3.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoRecordset3.EOF = False Then
                wkRecordset.Fields("Pname") = adoRecordset3.Fields("Pname")
            End If
            adoRecordset3.Close
            
            wkRecordset.Update
            
            adoRecordset2("Wdiv") = CHECK_REPORT_ON
            adoRecordset2.Update
            
            adoRecordset2.MoveNext
            intLineCount = intLineCount + 1
        Loop
        adoRecordset2.Close
        
'        If intLineCount <= PAGE_MAX_ROW Then
'            '空行作成
'            For intIndex1 = intLineCount To PAGE_MAX_ROW
'                wkRecordset.AddNew
'                wkRecordset.Fields("Ocode") = adoRecordset1.Fields("Ocode")
'                wkRecordset.Fields("Odate") = adoRecordset1.Fields("Odate")
'                wkRecordset.Fields("Hnum") = adoRecordset1.Fields("Hnum")
'                wkRecordset.Fields("Line") = intIndex1
'                wkRecordset.Update
'            Next intIndex1
'        End If
        adoRecordset1.MoveNext
    Loop
        
    adoRecordset1.Close
    wkRecordset.Requery
    wkRecordset.Close
    
    g_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    MakeWork = True
    
    Exit Function
    
MakeWork_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    Screen.MousePointer = vbDefault
    Call MsgBox("ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork_Err")
    
End Function


