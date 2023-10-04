VERSION 5.00
Begin VB.Form frmTorikesi 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "入金取消"
   ClientHeight    =   3750
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   8595
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "入金取消"
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
      Left            =   6060
      TabIndex        =   1
      Top             =   3240
      Width           =   1275
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
      Left            =   7440
      TabIndex        =   0
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "取消するデータを選択して入金取消ボタンをクリックしてください。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   34
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7755
   End
End
Attribute VB_Name = "frmTorikesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public g_intBocode As Integer
Public g_blnTorikesizumi As Boolean

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdExecute_Click()
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strOdate As String
    Dim strRdate As String
    
    On Error GoTo cmdExecute_Click_Err

    If List1.ListIndex < 0 Then
        Call MsgBox("データを選択してください", vbOKOnly + vbInformation, "")
        Exit Sub
    End If

    strOdate = Mid(List1.List(List1.ListIndex), 5, 10)
    strRdate = Mid(List1.List(List1.ListIndex), 20, 10)

'********** 入金取り消し処理 **********
                
    '入金データ削除
    strSQL = "DELETE FROM DT060" & _
             " WHERE Odate = '" & strOdate & "'" & _
             " AND Bcode = " & g_intBocode & _
             " AND Rdate = '" & strRdate & "'"
    g_clsAdoSQL.Connection.Execute strSQL
    
    '買主精算データ
    strSQL = "SELECT * FROM DT041" & _
             " WHERE Odate = '" & strOdate & "'" & _
             " AND Bcode = " & g_intBocode & _
             " ORDER BY Bcode,Odate,Num"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
    Do While Not adoRecordset1.EOF
        adoRecordset1.Fields("Rdiv") = PAYMENT_OFF
        adoRecordset1.Fields("R") = 0
        adoRecordset1.Fields("Rdate") = ""
        adoRecordset1.Update
    
        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close
    
    '買主精算データ(累積)
    strSQL = "SELECT * FROM RT041" & _
             " WHERE Odate = '" & strOdate & "'" & _
             " AND Bcode = " & g_intBocode & _
             " ORDER BY Bcode,Odate,Num"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
    Do While Not adoRecordset1.EOF
        adoRecordset1.Fields("Rdiv") = PAYMENT_OFF
        adoRecordset1.Fields("R") = 0
        adoRecordset1.Fields("Rdate") = ""
        adoRecordset1.Update
    
        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close
    
    g_blnTorikesizumi = True
    Call FieldRefresh
    
    Exit Sub
    
cmdExecute_Click_Err:

    Call MsgBox("入金取消クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    g_blnTorikesizumi = False
    Call FieldRefresh
    
    Exit Sub
    
Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    
End Sub

Private Sub FieldRefresh()
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim curNyukin As Currency
    Dim strBuff As String
    
    On Error GoTo FieldRefresh_Err

    'データオープン
    strSQL = "SELECT * FROM DT060" & _
             " WHERE Bcode = " & g_intBocode & _
             " ORDER BY Bcode,Odate DESC,Rdate DESC"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    
    List1.Clear
    Do While Not adoRecordset1.EOF
        strBuff = "開催日:" & adoRecordset1.Fields("Odate") & "　"
        strBuff = strBuff & "入金日:" & adoRecordset1.Fields("Rdate") & "　"
        
        Select Case adoRecordset1.Fields("R")
            Case PAYMENT_DIV_CASH:
                strBuff = strBuff & "[ 現　金 ]"
            Case PAYMENT_DIV_CHECK:
                strBuff = strBuff & "[小 切 手]"
            Case PAYMENT_DIV_TRANSFER:
                strBuff = strBuff & "[銀行振込]"
            Case Else
                strBuff = strBuff & "[ 現　金 ]"
        End Select
                
        curNyukin = CCur(adoRecordset1.Fields("Ptotal"))
                
        strBuff = strBuff & "　入金額: " & Format$(curNyukin, "#,##0")
                        
        List1.AddItem strBuff
                        
        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close
    
    Exit Sub
    
FieldRefresh_Err:

    Call MsgBox("データ表示エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldRefresh_Err")
    
End Sub


