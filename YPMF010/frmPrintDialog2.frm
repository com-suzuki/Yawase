VERSION 5.00
Begin VB.Form frmPrintDialog2 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "���x�����s"
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
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CheckBox chkFlg 
      Caption         =   "�ύX���̂ݔ��s"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Value           =   1  '����
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���x�����s���܂����H"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
        .Name = "�o�i�`�["
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_YPMF010"
        .Caption = "�o�i�`�["
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
    Call MsgBox("���s�N���b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Form_KeyDown_Err

    '���^�[���L�[�Ŏ��̃R���g���[���փt�H�[�J�X�ړ�
    If KeyCode = vbKeyReturn And Shift <> vbShiftMask Then
        KeyCode = 0
        Global_SendKeys Me, VK_TAB
        Exit Sub
    End If
    
    Exit Sub

Form_KeyDown_Err:

    Call MsgBox("�t�H�[���L�[�_�E�����G���[�I�I" _
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

    Call MsgBox("�t�H�[�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "���̓`�F�b�N")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("���̓`�F�b�N�G���[�I�I" _
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
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF010"
    g_clsAdoAccess.Connection.Execute strSQL
    
    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF010"
    wkRecordset1.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '�f�[�^�I�[�v��
    strSQL = "SELECT * FROM DT010" & _
             " WHERE Odate = '" & frmYpmf010.lblOdate.Caption & "'" & _
             " AND Pnum BETWEEN " & m_intPnum & " AND " & m_intPnum
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        intLineCount = 1
        
        '�f�[�^�I�[�v��
        strSQL = "SELECT * FROM vw_YPMF010" & _
                 " WHERE Odate = '" & adoRecordset1.Fields("Odate") & "'" & _
                 " AND Pnum = " & adoRecordset1.Fields("Pnum")
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset2.EOF
            blnAddnewflg = True
        
            '�ύX���݂̂̏ꍇ
            If chkFlg.Value = 1 Then
                blnAddnewflg = False
            
                '���[�N�I�[�v��
                strSQL = "SELECT * FROM WK_YPMF011" & _
                        " WHERE Odate = '" & frmYpmf010.lblOdate.Caption & "'" & _
                        " AND Pnum = " & m_intPnum & _
                        " AND Line = " & intLineCount
                wkRecordset2.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockReadOnly
                If wkRecordset2.EOF = False Then
                    '�f�[�^�����݂�����ύX�����Ƃ݂Ȃ�
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
    wkRecordset1.Requery     '�o�O�h�~
    wkRecordset1.Close
    
    '���[�N�̃f�[�^���݃`�F�b�N
    strSQL = "SELECT * FROM WK_YPMF010"
    wkRecordset1.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    If wkRecordset1.EOF = True Then
        wkRecordset1.Close
        Screen.MousePointer = vbDefault
        Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "")
        Exit Function
    End If
    wkRecordset1.Requery     '�o�O�h�~
    wkRecordset1.Close
    
    Screen.MousePointer = vbDefault
    
    MakeWork = True
    
    Exit Function
    
MakeWork_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("���[�N�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork_Err")
    
End Function


