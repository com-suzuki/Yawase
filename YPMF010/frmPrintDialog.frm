VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "Cscapt32.ocx"
Begin VB.Form frmPrintDialog 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "���"
   ClientHeight    =   1170
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5580
   Icon            =   "frmPrintDialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame Frame1 
      Caption         =   "���o����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4035
      Begin imText6Ctl.imText txtPnum 
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
            Name            =   "�l�r �o�S�V�b�N"
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
         Caption         =   "��t�ԍ�"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
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
      Begin imText6Ctl.imText txtPnum 
         Height          =   360
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   635
         Caption         =   "frmPrintDialog.frx":00DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
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
         Caption         =   "�`"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   420
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "��ݾ�"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   660
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
      Left            =   4200
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
    
    Dim objRpt As New rptYpmf010
    
    On Error GoTo cmdExecute_Click_Err
    
    If DoValidationChecks() = False Then Exit Sub
    If MakeWork() = False Then Exit Sub
    
    If m_blnAutoPrint = False Then cmdCancel.SetFocus
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "�o�i��t�m�F�\"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_YPMF010"
        .Caption = "�o�i��t�m�F�\"
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

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo Form_Load_Err

    txtPnum(0).Text = ""
    txtPnum(1).Text = ""
    
    If Trim(frmYpmf010.m_strLastPnum) = "" Then
        With adoRecordset1
            strSQL = "SELECT * FROM DT010" & _
                 " WHERE Odate = '" & frmYpmf010.lblOdate.Caption & "'" & _
                 " ORDER BY Pnum DESC"
            .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If Not .EOF Then
                txtPnum(0).Text = .Fields("Pnum")
                txtPnum(1).Text = .Fields("Pnum")
            End If
        End With
    Else
        txtPnum(0).Text = frmYpmf010.m_strLastPnum
        txtPnum(1).Text = frmYpmf010.m_strLastPnum
    End If
    
    If m_blnAutoPrint = True Then
        Call cmdExecute_Click
        Unload Me
    End If
    
    Exit Sub
    
Form_Load_Err:

    Call MsgBox("�t�H�[�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtPnum(0).Text) = "" Then
        strErrMsg = "��t�ԍ�����͂��Ă��������B"
        txtPnum(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtPnum(1).Text) = "" Then
        strErrMsg = "��t�ԍ�����͂��Ă��������B"
        txtPnum(1).SetFocus
        GoTo ErrorTrap:
    End If
    
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

Private Sub txtPnum_GotFocus(Index As Integer)

    txtPnum(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtPnum_LostFocus(Index As Integer)

    txtPnum(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Function MakeWork() As Boolean
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim intLineCount As Integer
    Dim intIndex1 As Integer
    
    Const PAGE_MAX_ROW = 20
    
    On Error GoTo MakeWork_Err
    
    MakeWork = False
    
    Screen.MousePointer = vbHourglass
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF010"
    g_clsAdoAccess.Connection.Execute strSQL
    
    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF010"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '�f�[�^�I�[�v��
    strSQL = "SELECT * FROM DT010" & _
             " WHERE Odate = '" & frmYpmf010.lblOdate.Caption & "'" & _
             " AND Pnum BETWEEN " & txtPnum(0).Text & " AND " & txtPnum(1).Text
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        intLineCount = 1
        
        '�f�[�^�I�[�v��
        strSQL = "SELECT * FROM vw_YPMF010" & _
                 " WHERE Odate = '" & adoRecordset1.Fields("Odate") & "'" & _
                 " AND Pnum = " & adoRecordset1.Fields("Pnum")
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset2.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Odate") = adoRecordset2.Fields("Odate")
            wkRecordset.Fields("Pnum") = adoRecordset2.Fields("Pnum")
            wkRecordset.Fields("Line") = intLineCount
            wkRecordset.Fields("Icode") = adoRecordset2.Fields("Icode")
            wkRecordset.Fields("Iname") = adoRecordset2.Fields("Iname")
            wkRecordset.Fields("Qty") = adoRecordset2.Fields("Qty")
            wkRecordset.Fields("Price1") = adoRecordset2.Fields("Price1")
            wkRecordset.Fields("Price2") = adoRecordset2.Fields("Price2")
            wkRecordset.Fields("Price") = adoRecordset2.Fields("Price")
            wkRecordset.Fields("Bcode") = adoRecordset2.Fields("Bcode")
            wkRecordset.Fields("Scode") = adoRecordset2.Fields("Scode")
            wkRecordset.Fields("Sname") = adoRecordset2.Fields("Sname")
            wkRecordset.Fields("Addres") = adoRecordset2.Fields("Addres")
            wkRecordset.Fields("Tel") = adoRecordset2.Fields("Tel")
            wkRecordset.Fields("Div") = adoRecordset2.Fields("Div")
            wkRecordset.Fields("Soukin") = adoRecordset1.Fields("Soukin")
            wkRecordset.Update
            
            adoRecordset2.MoveNext
            intLineCount = intLineCount + 1
        Loop
        adoRecordset2.Close
        
'        If intLineCount <= PAGE_MAX_ROW Then
'            '��s�쐬
'            For intIndex1 = intLineCount To PAGE_MAX_ROW
'                wkRecordset.AddNew
'                wkRecordset.Fields("Odate") = adoRecordset1.Fields("Odate")
'                wkRecordset.Fields("Pnum") = adoRecordset1.Fields("Pnum")
'                wkRecordset.Fields("Line") = intIndex1
'                wkRecordset.Update
'            Next intIndex1
'        End If
        adoRecordset1.MoveNext
    Loop
        
    adoRecordset1.Close
    wkRecordset.Requery     '�o�O�h�~
    wkRecordset.Close
    
    Screen.MousePointer = vbDefault
    
    MakeWork = True
    
    Exit Function
    
MakeWork_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("���[�N�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork_Err")
    
End Function


