VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmPrintDialog2 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�o�i�ғ`�[���"
   ClientHeight    =   2025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6510
   Icon            =   "frmPrintDialog2frm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6510
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
      Height          =   1875
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4995
      Begin VB.Frame Frame2 
         BorderStyle     =   0  '�Ȃ�
         Height          =   615
         Left            =   1620
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   3315
         Begin VB.OptionButton optTaisyou 
            Caption         =   "�o�i��"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
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
            TabIndex        =   13
            Top             =   120
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optTaisyou 
            Caption         =   "���@��"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   12
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.OptionButton optTanka 
         Caption         =   "����P��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3300
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optTanka 
         Caption         =   "�d���P��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
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
         TabIndex        =   8
         Top             =   1320
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin imText6Ctl.imText txtOnum 
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   635
         Caption         =   "frmPrintDialog2frm.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog2frm.frx":007A
         Key             =   "frmPrintDialog2frm.frx":0098
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
         Caption         =   "�����ԍ�"
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
      Begin imText6Ctl.imText txtOnum 
         Height          =   360
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   635
         Caption         =   "frmPrintDialog2frm.frx":00DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog2frm.frx":014A
         Key             =   "frmPrintDialog2frm.frx":0168
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
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "����P��"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "����Ώ�"
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
      Left            =   5160
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
      Left            =   5160
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmPrintDialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objArPrint As New clsArPrint
Public m_blnAutoPrint As Boolean

Private Sub cmdExecute_Click()
    
    On Error GoTo cmdExecute_Click_Err
    
    If optTaisyou(0).Value = True Then
        If PrintSyukkasya() = False Then Exit Sub
    ElseIf optTaisyou(1).Value = True Then
        If PrintKainusi() = False Then Exit Sub
    End If
    
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

    txtOnum(0).Text = ""
    txtOnum(1).Text = ""
    
    With adoRecordset1
        strSQL = "SELECT Onum FROM DT030" & _
             " WHERE Odate = '" & g_strOdate & "'" & _
             " ORDER BY Onum"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            txtOnum(0).Text = .Fields("Onum")
            .MoveLast
            txtOnum(1).Text = .Fields("Onum")
        End If
        .Close
    End With
    
    Exit Sub
    
Form_Load_Err:

    Call MsgBox("�t�H�[�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtOnum(0).Text) = "" Then
        strErrMsg = "�����ԍ�����͂��Ă��������B"
        txtOnum(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtOnum(1).Text) = "" Then
        strErrMsg = "�����ԍ�����͂��Ă��������B"
        txtOnum(1).SetFocus
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

Private Sub optTaisyou_Click(Index As Integer)

    If optTaisyou(Index).Value = True Then
        optTanka(Index).Value = True
    End If

End Sub

Private Sub txtOnum_GotFocus(Index As Integer)

    txtOnum(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtOnum_LostFocus(Index As Integer)

    txtOnum(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Function MakeWork() As Boolean
    
    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim intLineCount As Integer
    Dim intIndex1 As Integer
    
    Dim curSkeep As Currency                '�o�i�҈ێ��Ǘ���
    Dim curSrate As Currency                '�o�i�Ҏ萔����
    Dim intSfraction As Integer             '�o�i�Ғ[������
    Dim curSRounding As Currency            '�o�i�ҊۂߒP��
    Dim curTaxRate As Currency              '����ŗ�
    
    Dim curTotal As Currency
    Dim curCharge As Currency
    Dim curTax As Currency
    Dim curKeep As Currency
    Dim curGTotal As Currency
    
    On Error GoTo MakeWork_Err
    
    MakeWork = False
    
    Screen.MousePointer = vbHourglass
    
    '������
    curSkeep = 0        '�o�i�҈ێ��Ǘ���
    intSfraction = 0    '�o�i�Ғ[������
    curSrate = 0        '�o�i�Ҏ萔����
    curSRounding = 0    '�o�i�ҊۂߒP��
    curTaxRate = 0      '����ŗ�
    
    '�ݒ�}�X�^�I�[�v��
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        If Not IsNull(adoMT010.Fields("Skeep")) Then curSkeep = adoMT010.Fields("Skeep")
        If Not IsNull(adoMT010.Fields("Sfraction")) Then intSfraction = adoMT010.Fields("Sfraction")
        If Not IsNull(adoMT010.Fields("Srate")) Then curSrate = adoMT010.Fields("Srate")
        If Not IsNull(adoMT010.Fields("SRounding")) Then curSRounding = adoMT010.Fields("SRounding")
    End If
    adoMT010.Close
    
    '����ŗ��擾
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, g_strOdate)
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF031"
    g_clsAdoAccess.Connection.Execute strSQL
    
    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF031"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '�f�[�^�I�[�v��
    strSQL = "SELECT * FROM DT030" & _
             " WHERE Odate = '" & g_strOdate & "'" & _
             " AND Onum BETWEEN " & txtOnum(0).Text & " AND " & txtOnum(1).Text & _
             " ORDER BY Odate,Onum"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        curTotal = 0
        curCharge = 0
        curTax = 0
        curKeep = 0
        curGTotal = 0
        
        intLineCount = 1
        
        '�f�[�^�I�[�v��
        strSQL = "SELECT * FROM DT031" & _
                 " WHERE Odate = '" & adoRecordset1.Fields("Odate") & "'" & _
                 " AND Onum = " & adoRecordset1.Fields("Onum") & _
                 " ORDER BY Odate,Onum,Line"
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset2.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Key1") = adoRecordset1.Fields("Onum")
            wkRecordset.Fields("Key2") = adoRecordset1.Fields("Onum")
            wkRecordset.Fields("Num") = 1
            wkRecordset.Fields("Div") = "B"
            wkRecordset.Fields("Odate") = adoRecordset2.Fields("Odate")
            wkRecordset.Fields("Pnum") = adoRecordset1.Fields("Onum")
            wkRecordset.Fields("Scode") = adoRecordset1.Fields("Scode")
            wkRecordset.Fields("Sname") = Trim(adoRecordset1.Fields("Sname")) & "�@�l"
            wkRecordset.Fields("Line") = intLineCount
            wkRecordset.Fields("Icode") = adoRecordset2.Fields("Icode")
            wkRecordset.Fields("Iname") = adoRecordset2.Fields("Iname")
            wkRecordset.Fields("Qty") = adoRecordset2.Fields("Qty")
            If optTanka(0).Value = True Then
                '���z�����ʁ~�d���P��
                wkRecordset.Fields("Price1") = CCur(adoRecordset2.Fields("Qty")) * CCur(adoRecordset2.Fields("Price2"))
            ElseIf optTanka(1).Value = True Then
                '���z�����ʁ~����P��
                wkRecordset.Fields("Price1") = CCur(adoRecordset2.Fields("Qty")) * CCur(adoRecordset2.Fields("Price1"))
            End If
            wkRecordset.Fields("Bcode") = adoRecordset2.Fields("Bcode")
'            wkRecordset.Fields("Bname") = adoRecordset2.Fields("Bname")
'            wkRecordset.Fields("Total") = 0
'            wkRecordset.Fields("Charge") = 0
'            wkRecordset.Fields("Tax") = 0
'            wkRecordset.Fields("Keep") = 0
'            wkRecordset.Fields("GTotal") = 0
            wkRecordset.Fields("Itime") = Format(Now(), "nn:ss")
            wkRecordset.Fields("Pcode") = g_strPcode
            wkRecordset.Fields("Pname") = g_strPname
            wkRecordset.Fields("Idiv") = 0
            wkRecordset.Fields("Result") = 0
            wkRecordset.Fields("Ocode") = Format(adoRecordset1.Fields("Onum"), "0000")
            wkRecordset.Fields("RePrint") = 0
            
            wkRecordset.Update
            
            curTotal = curTotal + CCur(wkRecordset.Fields("Price1"))
            
            adoRecordset2.MoveNext
            intLineCount = intLineCount + 1
        Loop
        adoRecordset2.Close
        
        '�萔��
        If IsNull(adoRecordset1.Fields("ChargeDiv")) = False And adoRecordset1.Fields("ChargeDiv") = 1 Then
            curCharge = curTotal * curSrate / 100
            '�؂�̂�
            If IsNull(adoRecordset1.Fields("FixDiv")) = False And adoRecordset1.Fields("FixDiv") = 1 Then
                curCharge = Global_Rounding(curCharge, intSfraction, curSRounding)
            Else
                curCharge = Global_Rounding(curCharge, intSfraction, 1)
            End If
        End If
        '�ێ��Ǘ���
        If IsNull(adoRecordset1.Fields("KeepDiv")) = False And adoRecordset1.Fields("KeepDiv") = 1 Then
            curKeep = curSkeep
        End If
        '����Ōv�Z
        If IsNull(adoRecordset1.Fields("TaxDiv")) = False And adoRecordset1.Fields("TaxDiv") = 1 Then
            '�؂�̂�
            If IsNull(adoRecordset1.Fields("FixDiv")) = False And adoRecordset1.Fields("FixDiv") = 1 Then
                '202308 �C���{�C�X�Ή�
                curTax = Global_Get_Tax(curTotal - curCharge - curKeep, curTaxRate, intSfraction, 1)

            Else
                '202308 �C���{�C�X�Ή�
                curTax = Global_Get_Tax(curTotal - curCharge - curKeep, curTaxRate, intSfraction, 1)
            End If
        End If
        
        '�����v�����v�|�萔���|�ێ��Ǘ���
        curGTotal = curTotal - curCharge - curKeep + curTax
        '���[�N�̍X�V
        '202308 �C���{�C�X�Ή��@�o�i�ғo�^�ԍ��X�V�@WK_YPMF031.Invoice = '" & Trim(adoRecordset1.Fields("Addres")) & "'"
        strSQL = "UPDATE WK_YPMF031" & _
                 " SET WK_YPMF031.Total = " & curTotal & "," & _
                 " WK_YPMF031.Charge = " & curCharge & "," & _
                 " WK_YPMF031.Tax = " & curTax & "," & _
                 " WK_YPMF031.Keep = " & curKeep & "," & _
                 " WK_YPMF031.GTotal = " & curGTotal & "," & _
                 " WK_YPMF031.Bname = '" & Trim(adoRecordset1.Fields("Addres")) & "'" & _
                 " WHERE WK_YPMF031.Odate = '" & adoRecordset1.Fields("Odate") & "'" & _
                 " AND WK_YPMF031.Key1 = '" & adoRecordset1.Fields("Onum") & "'"

        g_clsAdoAccess.Connection.Execute strSQL


        
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

Private Function PrintSyukkasya() As Boolean
    
    Dim objRpt As New rptYpmf031
    
    On Error GoTo PrintSyukkasya_Err
    
    If DoValidationChecks() = False Then Exit Function
    If MakeWork() = False Then Exit Function
    
    If m_blnAutoPrint = False Then cmdCancel.SetFocus
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "�x���`�["
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_Ypmf031"
        .Caption = "�x���`�["
        If m_blnAutoPrint = True Then
            If .PrintActiveReport(1) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        Else
            If .PrintActiveReport(0) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    End With
    
    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
PrintSyukkasya_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("�o�׎҈���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "PrintSyukkasya_Err")
    
End Function

Private Function PrintKainusi() As Boolean
    
    Dim objRpt As New rptYpmf032
    
    On Error GoTo PrintKainusi_Err
    
    If DoValidationChecks() = False Then Exit Function
    If MakeWork2() = False Then Exit Function
    
    If m_blnAutoPrint = False Then cmdCancel.SetFocus
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "����`�[�i�����j"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_Ypmf032"
        .Caption = "����`�[�i�����j"
        If m_blnAutoPrint = True Then
            If .PrintActiveReport(1) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        Else
            If .PrintActiveReport(0) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    End With
    
    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
PrintKainusi_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("�������G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "PrintKainusi_Err")
    
End Function

'�ځ@�I�@�@�F����p���[�N�쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�V�^�O�W
'�X�V�����@�F
'
Private Function MakeWork2() As Boolean

    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim adoDT011 As New ADODB.Recordset
    Dim adoDT011M As New ADODB.Recordset
    Dim adoDT031 As New ADODB.Recordset
    Dim adoDT031M As New ADODB.Recordset
    Dim adoDT041 As New ADODB.Recordset
    Dim adoDT041TEMP As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim wkRecordsetTemp As New ADODB.Recordset
    Dim strBuff1 As String
    Dim blnAddNewFlg As Boolean
    
    Dim intLine As Integer                  '�s�ԍ�
    Dim lngCount As Long                    '���R�[�h����
    Dim curBkeep As Currency                '����ێ��Ǘ���(�W��)
    Dim curBkeepCurrent As Currency         '����ێ��Ǘ���(����)
    Dim intBfraction As Integer             '����[������
    Dim intNum As Integer                   '��
    Dim curTaxRate As Currency              '����ŗ�
    Dim intR As Integer                     '�������
    Dim curBRounding As Currency            '����ۂߒP��
    Dim strMemo As String                   '�`�[�̃���
    Dim strOdateNum As String

    Dim curPriceTotal As Currency
    Dim curTax As Currency
    Dim curGTotal As Currency

    On Error GoTo MakeWork2_Err
    
    MakeWork2 = False
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
'********** �������� **********
    
    '������
    lngCount = 0        '���R�[�h����
    curBkeep = 0        '����ێ��Ǘ���
    intBfraction = 0    '����[������
    curTaxRate = 0      '����ŗ�
    intR = 0            '�������
    curBRounding = 0    '����ۂߒP��
    strMemo = ""
    
    '�ݒ�}�X�^�I�[�v��
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        If Not IsNull(adoMT010.Fields("Bkeep")) Then curBkeep = adoMT010.Fields("Bkeep")
        If Not IsNull(adoMT010.Fields("Bfraction")) Then intBfraction = adoMT010.Fields("Bfraction")
        If Not IsNull(adoMT010.Fields("BRounding")) Then curBRounding = adoMT010.Fields("BRounding")
        If Not IsNull(adoMT010.Fields("Memo")) Then strMemo = adoMT010.Fields("Memo")
    End If
    adoMT010.Close
    
    '����ŗ��擾
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, g_strOdate)
    
'********** ���[�N **********
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF032"
    g_clsAdoAccess.Connection.Execute strSQL

    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF032"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
'********** �����f�[�^ **********
    
    '�����f�[�^�I�[�v��
    strSQL = "SELECT Bcode FROM DT031" & _
             " WHERE Bcode IS NOT NULL" & _
             " GROUP BY Bcode" & _
             " ORDER BY Bcode"
    adoDT031.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    
    lngCount = 0
    Do While Not adoDT031.EOF
        intLine = 1     '�s�ԍ�
        curPriceTotal = 0
        
        '�������׃f�[�^�I�[�v��
        strSQL = "SELECT * FROM DT031" & _
                 " WHERE Bcode = " & adoDT031.Fields("Bcode") & _
                 " ORDER BY Bcode,Onum"
        adoDT031M.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT031M.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Key1") = Format(adoDT031M.Fields("Bcode"), "0000")
            wkRecordset.Fields("Key2") = Format(adoDT031M.Fields("Bcode"), "0000")
            wkRecordset.Fields("Num") = 1
            wkRecordset.Fields("Div") = "B"
            wkRecordset.Fields("Odate") = g_strOdate
            wkRecordset.Fields("Bcode") = adoDT031M.Fields("Bcode")
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, adoDT031M.Fields("Bcode"), g_strOdate, strBuff1)
            wkRecordset.Fields("Icode") = adoDT031M.Fields("Icode")
            wkRecordset.Fields("Iname") = adoDT031M.Fields("Iname")
            wkRecordset.Fields("Qty") = adoDT031M.Fields("Qty")
            If optTanka(0).Value = True Then
                wkRecordset.Fields("Price1") = CCur(adoDT031M.Fields("Qty")) * CCur(adoDT031M.Fields("Price2"))
            ElseIf optTanka(1).Value = True Then
                wkRecordset.Fields("Price1") = adoDT031M.Fields("Price")
            End If
            wkRecordset.Fields("Pnum") = adoDT031M.Fields("Onum")
            wkRecordset.Fields("Sname") = ""
'            wkRecordset.Fields("Total") = 0
'            wkRecordset.Fields("Tax") = 0
            wkRecordset.Fields("Keep") = curBkeep
'            wkRecordset.Fields("GTotal") = 0
            wkRecordset.Fields("Itime") = Format(Now(), "hh��mm��")
            wkRecordset.Fields("Pcode") = g_strPcode
            wkRecordset.Fields("Pname") = g_strPname
            wkRecordset.Fields("Idiv") = 0
            wkRecordset.Fields("Ocode") = right$(Format(adoDT031M.Fields("Onum"), "0000"), 4) & "*"
            wkRecordset.Fields("Memo") = strMemo
            wkRecordset.Update
        
            curPriceTotal = curPriceTotal + CCur(adoDT031M.Fields("Price"))
        
            adoDT031M.MoveNext
            intLine = intLine + 1
        Loop
        adoDT031M.Close
        
        '�����v�v�Z
        curGTotal = curPriceTotal - curBkeep
        curTax = Global_Get_Tax(curPriceTotal, curTaxRate, intBfraction, curBRounding)
        
        
        '���[�N�̍��v�z�Ȃǂ��X�V
        strSQL = "UPDATE WK_YPMF032" & _
                 " SET WK_YPMF032.Total = " & curPriceTotal & "," & _
                 " WK_YPMF032.Tax = " & curTax & "," & _
                 " WK_YPMF032.Keep = " & curBkeep & "," & _
                 " WK_YPMF032.GTotal = " & curGTotal & _
                 " WHERE WK_YPMF032.Odate = '" & g_strOdate & "'" & _
                 " AND WK_YPMF032.Bcode = " & adoDT031.Fields("Bcode")
        g_clsAdoAccess.Connection.Execute strSQL
        
        adoDT031.MoveNext
        lngCount = lngCount + 1
    Loop
    adoDT031.Close
    
    wkRecordset.Requery
    wkRecordset.Close
    
    g_clsAdoSQL.Connection.CommitTrans
    
    If lngCount = 0 Then
        Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "")
        GoTo MakeWork2_Exit:
    End If
    
    MakeWork2 = True
    
MakeWork2_Exit:
    
    Screen.MousePointer = vbDefault
    
    Exit Function

MakeWork2_Cancel:

    g_clsAdoSQL.Connection.RollbackTrans
    GoTo MakeWork2_Exit:

MakeWork2_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    MakeWork2 = False
    Call MsgBox("������[�N�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork2_Err")
    GoTo MakeWork2_Exit:

End Function

