VERSION 5.00
Begin VB.Form frmPrintDialog2 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ラベル発行"
   ClientHeight    =   1170
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4605
   Icon            =   "frmPrintDialog2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CheckBox chkFlg 
      Caption         =   "変更分のみ発行"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Value           =   1  'ﾁｪｯｸ
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "閉じる"
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
      Left            =   3240
      TabIndex        =   1
      Top             =   600
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
      Left            =   3240
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "ラベル発行しますか？"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   2895
   End
End
Attribute VB_Name = "frmPrintDialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objArPrint As New clsArPrint
Public m_intPnum As Integer
Public m_intMode As Integer

Private Sub cmdExecute_Click()

    Dim objRpt As New rptYpmf010B
    Dim objArPrint As New clsArPrint
    
    On Error GoTo cmdExecute_Click_Err
    
    If MakeWork() = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "出品伝票"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_YPMF010"
        .Caption = "出品伝票"
        If .PrintActiveReport(0) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End With
    
    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    Unload Me
    
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

    On Error GoTo Form_Load_Err
    
    If m_intMode = 0 Then
        chkFlg.Value = 0
        chkFlg.Visible = False
    ElseIf m_intMode = 0 Then
        chkFlg.Value = 1
        chkFlg.Visible = True
    End If
    
    Exit Sub
    
Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

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

Private Function MakeWork() As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim wkRecordset1 As New ADODB.Recordset
    Dim wkRecordset2 As New ADODB.Recordset
    Dim intLineCount As Integer
    Dim intIndex1 As Integer
    Dim blnAddnewflg As Boolean
    
    Const PAGE_MAX_ROW = 20
    
    On Error GoTo MakeWork_Err
    
    MakeWork = False
    
    Screen.MousePointer = vbHourglass
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF010"
    g_clsAdoAccess.Connection.Execute strSQL
    
    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF010"
    wkRecordset1.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    'データオープン
    strSQL = "SELECT * FROM DT010" & _
             " WHERE Odate = '" & frmYpmf010.lblOdate.Caption & "'" & _
             " AND Pnum BETWEEN " & m_intPnum & " AND " & m_intPnum
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        intLineCount = 1
        
        'データオープン
        strSQL = "SELECT * FROM vw_YPMF010" & _
                 " WHERE Odate = '" & adoRecordset1.Fields("Odate") & "'" & _
                 " AND Pnum = " & adoRecordset1.Fields("Pnum")
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset2.EOF
            blnAddnewflg = True
        
            '変更分のみの場合
            If chkFlg.Value = 1 Then
                blnAddnewflg = False
            
                'ワークオープン
                strSQL = "SELECT * FROM WK_YPMF011" & _
                        " WHERE Odate = '" & frmYpmf010.lblOdate.Caption & "'" & _
                        " AND Pnum = " & m_intPnum & _
                        " AND Line = " & intLineCount
                wkRecordset2.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockReadOnly
                If wkRecordset2.EOF = False Then
                    'データが存在したら変更したとみなす
                    blnAddnewflg = True
                End If
                wkRecordset2.Close
            End If
            
            If blnAddnewflg = True Then
                wkRecordset1.AddNew
                wkRecordset1.Fields("Odate") = adoRecordset2.Fields("Odate")
                wkRecordset1.Fields("Pnum") = adoRecordset2.Fields("Pnum")
                wkRecordset1.Fields("Line") = intLineCount
                wkRecordset1.Fields("Icode") = adoRecordset2.Fields("Icode")
                wkRecordset1.Fields("Iname") = adoRecordset2.Fields("Iname")
                wkRecordset1.Fields("Qty") = adoRecordset2.Fields("Qty")
                wkRecordset1.Fields("Price1") = adoRecordset2.Fields("Price1")
                wkRecordset1.Fields("Price2") = adoRecordset2.Fields("Price2")
                wkRecordset1.Fields("Price") = adoRecordset2.Fields("Price")
                wkRecordset1.Fields("Bcode") = adoRecordset2.Fields("Bcode")
                wkRecordset1.Fields("Scode") = adoRecordset2.Fields("Scode")
                wkRecordset1.Fields("Sname") = adoRecordset2.Fields("Sname")
                wkRecordset1.Fields("Addres") = adoRecordset2.Fields("Addres")
                wkRecordset1.Fields("Tel") = adoRecordset2.Fields("Tel")
                wkRecordset1.Fields("Div") = adoRecordset2.Fields("Div")
                wkRecordset1.Fields("Soukin") = adoRecordset1.Fields("Soukin")
                wkRecordset1.Update
            End If
            
            adoRecordset2.MoveNext
            intLineCount = intLineCount + 1
        Loop
        adoRecordset2.Close
        
        adoRecordset1.MoveNext
    Loop
        
    adoRecordset1.Close
    wkRecordset1.Requery     'バグ防止
    wkRecordset1.Close
    
    'ワークのデータ存在チェック
    strSQL = "SELECT * FROM WK_YPMF010"
    wkRecordset1.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    If wkRecordset1.EOF = True Then
        wkRecordset1.Close
        Screen.MousePointer = vbDefault
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
        Exit Function
    End If
    wkRecordset1.Requery     'バグ防止
    wkRecordset1.Close
    
    Screen.MousePointer = vbDefault
    
    MakeWork = True
    
    Exit Function
    
MakeWork_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork_Err")
    
End Function


