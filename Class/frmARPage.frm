VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmARPage 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�y�[�W�ݒ�"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "frmARPage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "��ݾ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   12
      Top             =   5100
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "O K"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   11
      Top             =   5100
      Width           =   1455
   End
   Begin VB.Frame fraMargin 
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   4395
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   6075
      Begin VB.TextBox txtRight 
         Alignment       =   1  '�E����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   1620
         Width           =   915
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  '�E����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   2
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtBottom 
         Alignment       =   1  '�E����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txtTop 
         Alignment       =   1  '�E����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   0
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   1800
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   1800
         TabIndex        =   21
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   1800
         TabIndex        =   20
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1800
         TabIndex        =   19
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "�E�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "���F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "���F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "��F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   420
         Width           =   435
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4875
      Left            =   60
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8599
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "�]�@��"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "�y�[�W"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraPage 
      BorderStyle     =   0  '�Ȃ�
      Height          =   4335
      Left            =   120
      TabIndex        =   23
      Top             =   540
      Width           =   5955
      Begin VB.Frame Frame4 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   60
         TabIndex        =   28
         Top             =   3420
         Width           =   4095
         Begin VB.ComboBox cboPaperSource 
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmARPage.frx":000C
            Left            =   120
            List            =   "frmARPage.frx":0013
            Style           =   2  '��ۯ���޳� ؽ�
            TabIndex        =   10
            Top             =   300
            Width           =   3855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "�v�����^"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   60
         TabIndex        =   27
         Top             =   1740
         Width           =   5835
         Begin VB.OptionButton optDeviceName2 
            Caption         =   "�w��̃v�����^"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   8
            Top             =   660
            Width           =   3315
         End
         Begin VB.OptionButton optDeviceName1 
            Caption         =   "�ʏ�g���v�����^"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   3315
         End
         Begin VB.ComboBox cboDeviceName 
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            Style           =   2  '��ۯ���޳� ؽ�
            TabIndex        =   9
            Top             =   1140
            Width           =   5475
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�p��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   60
         TabIndex        =   25
         Top             =   840
         Width           =   5835
         Begin VB.ComboBox cboPaperSize 
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1020
            Style           =   2  '��ۯ���޳� ؽ�
            TabIndex        =   6
            Top             =   240
            Width           =   4515
         End
         Begin VB.Label Label1 
            Caption         =   "�T�C�Y"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   180
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "����̌���"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   2775
         Begin VB.OptionButton optOrientation2 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1620
            TabIndex        =   5
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton optOrientation1 
            Caption         =   "�c"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   420
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "frmARPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_objArPrint As clsArPrint
Public m_blnCanselFlg As Boolean

'�v�����^�f�o�C�X�h���C�o�̔\�͂��擾����֐��̐錾
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal pDevice As String, ByVal pPort As String, ByVal fwCapability As Long, pOutput As Any, pDevMode As Any) As Long
'����ʒu����ʂ̈ʒu�Ƀ������u���b�N���ړ�����֐��̐錾
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'�萔�錾
Private Const DC_PAPERS = 2
Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12      '�v�����^�Ŏg�p�ł��鋋�����@�̖��O���擾���邱�Ƃ�����
Private Const DC_PAPERNAMES = 16    '�v�����^�Ŏg�p�ł���p���̖��O���擾����萔

Private Sub cboDeviceName_DropDown()

    Dim objPrinter As Printer
    
    cboDeviceName.Clear
    For Each objPrinter In Printers
        '�v�����^���̎擾
        cboDeviceName.AddItem objPrinter.DeviceName
    Next
    
End Sub

Private Sub cboPaperSize_DropDown()
    
    '�p���T�C�Y�̈ꗗ�\��
    Call ListPaperSizes

End Sub

Private Sub cboPaperSource_DropDown()
    
    Dim strBuff As String
    
    On Error GoTo cboPaperSource_DropDown_Err
    
    '�������@�̈ꗗ�\��
    strBuff = cboPaperSource.Text
    Call ListPaperSource
    
    '�R���{�{�b�N�X�̕\��
    If GetCboListName(cboPaperSource, strBuff) = False Then
        cboPaperSource.Text = cboPaperSource.List(0)
    End If

    Exit Sub
    
cboPaperSource_DropDown_Err:
    
    Call MsgBox("�������@�h���b�v�_�E�����G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "cboPaperSource_DropDown_Err")
    
End Sub

Private Sub cmdCancel_Click()

    m_blnCanselFlg = True
    Unload Me

End Sub

Private Sub cmdOk_Click()

    On Error GoTo cmdOk_Click_Err

    '���̓`�F�b�N
    If IsNumeric(txtTop.Text) = False Then
        fraMargin.ZOrder 0
        txtTop.SetFocus
        Call MsgBox("�������]������͂��Ă��������B", vbOKOnly + vbCritical, "���̓`�F�b�N")
        Exit Sub
    End If
    If IsNumeric(txtBottom.Text) = False Then
        fraMargin.ZOrder 0
        txtBottom.SetFocus
        Call MsgBox("�������]������͂��Ă��������B", vbOKOnly + vbCritical, "���̓`�F�b�N")
        Exit Sub
    End If
    If IsNumeric(txtLeft.Text) = False Then
        fraMargin.ZOrder 0
        txtLeft.SetFocus
        Call MsgBox("�������]������͂��Ă��������B", vbOKOnly + vbCritical, "���̓`�F�b�N")
        Exit Sub
    End If
    If IsNumeric(txtRight.Text) = False Then
        fraMargin.ZOrder 0
        txtRight.SetFocus
        Call MsgBox("�������]������͂��Ă��������B", vbOKOnly + vbCritical, "���̓`�F�b�N")
        Exit Sub
    End If
    If Trim(cboPaperSize.Text) = "" Then
        fraPage.ZOrder 0
        cboPaperSize.SetFocus
        Call MsgBox("�y�[�W�T�C�Y����͂��Ă��������B", vbOKOnly + vbCritical, "���̓`�F�b�N")
        Exit Sub
    End If
    If optDeviceName2.Value = True And Trim(cboDeviceName.Text) = "" Then
        fraPage.ZOrder 0
        cboPaperSize.SetFocus
        Call MsgBox("�v�����^����͂��Ă��������B", vbOKOnly + vbCritical, "���̓`�F�b�N")
        Exit Sub
    End If

    '�v���p�e�B�Z�b�g
    m_objArPrint.PrnPageTopMargin = txtTop.Text
    m_objArPrint.PrnPageBottomMargin = txtBottom.Text
    m_objArPrint.PrnPageLeftMargin = txtLeft.Text
    m_objArPrint.PrnPageRightMargin = txtRight.Text
    
    If optOrientation1.Value = True Then
        m_objArPrint.PrnOrientation = vbPRORPortrait
    ElseIf optOrientation2.Value = True Then
        m_objArPrint.PrnOrientation = vbPRORLandscape
    End If
    m_objArPrint.PrnPaperSize = GetPaperSize()
    If optDeviceName1.Value = True Then
        m_objArPrint.PrnDefaultPrinter = True
        m_objArPrint.PrnDeviceName = ""
    ElseIf optDeviceName2.Value = True Then
        m_objArPrint.PrnDefaultPrinter = False
        m_objArPrint.PrnDeviceName = cboDeviceName.Text
    End If
    If cboPaperSource.ListIndex = 0 Then
        m_objArPrint.PrnPaperSource = 0 '�����̏ꍇ�̓[���ɂ��Ă���
    Else
        m_objArPrint.PrnPaperSource = cboPaperSource.ItemData(cboPaperSource.ListIndex)
    End If
    
    '�t�H�[�������
    Unload Me
    m_blnCanselFlg = False
    
    '�v���r���[���e�����t���b�V��
    m_objArPrint.Refresh
    m_objArPrint.objReport.Restart
    m_objArPrint.objReport.Run
    frmARPreview.arv.ReportSource = m_objArPrint.objReport
    DoEvents
    
    Exit Sub
    
cmdOk_Click_Err:
    
    Call MsgBox("OK�{�^���N���b�N���G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdOk_Click_Err")
    
End Sub

Private Sub Form_Load()

    Dim intIndex1 As Integer

    On Error GoTo Form_Load_Err

    '�����ݒ�
    fraMargin.ZOrder 0
    cboDeviceName.Enabled = False

    '�����\��
    txtTop.Text = m_objArPrint.PrnPageTopMargin
    txtBottom.Text = m_objArPrint.PrnPageBottomMargin
    txtLeft.Text = m_objArPrint.PrnPageLeftMargin
    txtRight.Text = m_objArPrint.PrnPageRightMargin
    
    Select Case m_objArPrint.PrnOrientation
        Case vbPRORPortrait:
            optOrientation1.Value = True
        Case vbPRORLandscape:
            optOrientation2.Value = True
        Case Else
            optOrientation2.Value = True
    End Select
    If m_objArPrint.PrnDefaultPrinter = True Then
        optDeviceName1.Value = True
    Else
        optDeviceName2.Value = True
        '�R���{�{�b�N�X�쐬
        Call cboDeviceName_DropDown
        For intIndex1 = 0 To cboDeviceName.ListCount - 1
            If Trim(cboDeviceName.List(intIndex1)) = Trim(m_objArPrint.PrnDeviceName) Then
                cboDeviceName.Text = cboDeviceName.List(intIndex1)
                Exit For
            End If
        Next intIndex1
    End If
    
    '�p���T�C�Y
    Call GetPaperSizeName(m_objArPrint.PrnPaperSize)
    
    '�������@
    Call GetPaperSource(m_objArPrint.PrnPaperSource)
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("�t�H�[�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
 
End Sub

Private Sub optDeviceName1_Click()

    '�v�����^���̃N���A
    cboDeviceName.Clear
    cboDeviceName.Enabled = False
    
End Sub

Private Sub optDeviceName2_Click()

    cboDeviceName.Enabled = True

End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.Index
        Case 1:
            fraMargin.ZOrder 0
        Case 2:
            fraPage.ZOrder 0
    End Select
    
End Sub

Private Function GetPaperSizeName(intPaperSize As Integer) As Boolean
    
    Dim strPrinterDeviceName     As String
    Dim strPrinterPortName       As String
    Dim lngDeviceCapability      As Long
    Dim lngSupportedPapersNeeded As Long
    Dim intSupportedPapers()     As Integer
    Dim lngSupportedPapersCount  As Long
    Dim lngWin32apiResultCode    As Long
    
    On Error GoTo GetPaperSizeName_Err
    
    If intPaperSize = 0 Then
        Exit Function
    End If
    
    '�R���{�{�b�N�X�쐬
    If ListPaperSizes() = False Then Exit Function
    
    '�v�����^�f�o�C�X���ƃ|�[�g�����w�肷��
    If optDeviceName1.Value = True Then
        '�ʏ�g���v�����^�̏ꍇ
        strPrinterDeviceName = Printer.DeviceName
        strPrinterPortName = Printer.Port
    ElseIf optDeviceName2.Value = True Then
        If Trim(cboDeviceName.Text) = "" Then Exit Function
        '�w��̃v�����^�̏ꍇ
        Dim objPrinter As Printer
        
        For Each objPrinter In Printers
            If Trim(objPrinter.DeviceName) = Trim(cboDeviceName.Text) Then
                strPrinterDeviceName = objPrinter.DeviceName
                strPrinterPortName = objPrinter.Port
                Exit For
            End If
        Next
    End If
    
    '�₢���킹��\�͂��w��
    lngDeviceCapability = DC_PAPERS
    '�o�b�t�@�̕K�v�ȃT�C�Y���擾
    lngSupportedPapersNeeded = DeviceCapabilities(strPrinterDeviceName, _
                                                  strPrinterPortName, _
                                                  lngDeviceCapability, _
                                                  ByVal vbNullString, _
                                                  ByVal vbNullString)
    '�o�b�t�@���m��
    ReDim intSupportedPapers(lngSupportedPapersNeeded - 1)
    '�T�|�[�g����Ă���p���T�C�Y���擾
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, _
                                               strPrinterPortName, _
                                               lngDeviceCapability, _
                                               intSupportedPapers(0), _
                                               ByVal vbNullString)
    '�o�b�t�@�ɗp���T�C�Y�����݂����
    For lngSupportedPapersCount = 0 To lngSupportedPapersNeeded - 1
        '�T�|�[�g����Ă���p���T�C�Y���
        If intSupportedPapers(lngSupportedPapersCount) = intPaperSize Then
            '�p���T�C�Y���̃Z�b�g
            cboPaperSize.Text = cboPaperSize.List(lngSupportedPapersCount)
            Exit For
        End If
    Next lngSupportedPapersCount

    Exit Function
    
GetPaperSizeName_Err:

    GetPaperSizeName = ""
    Call MsgBox("�p���T�C�Y���擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "GetPaperSizeName_Err")
    
End Function

Private Function GetPaperSize() As Integer

    Dim strPrinterDeviceName     As String
    Dim strPrinterPortName       As String
    Dim lngDeviceCapability      As Long
    Dim lngSupportedPapersNeeded As Long
    Dim intSupportedPapers()     As Integer
    Dim lngSupportedPapersCount  As Long
    Dim lngWin32apiResultCode    As Long

    On Error GoTo GetPaperSize_Err
    
    GetPaperSize = 0
    
    '�v�����^�f�o�C�X���ƃ|�[�g�����w�肷��
    If optDeviceName1.Value = True Then
        '�ʏ�g���v�����^�̏ꍇ
        strPrinterDeviceName = Printer.DeviceName
        strPrinterPortName = Printer.Port
    ElseIf optDeviceName2.Value = True Then
        If Trim(cboDeviceName.Text) = "" Then Exit Function
        '�w��̃v�����^�̏ꍇ
        Dim objPrinter As Printer
        
        For Each objPrinter In Printers
            If Trim(objPrinter.DeviceName) = Trim(cboDeviceName.Text) Then
                strPrinterDeviceName = objPrinter.DeviceName
                strPrinterPortName = objPrinter.Port
                Exit For
            End If
        Next
    End If
    
    '�₢���킹��\�͂��w��
    lngDeviceCapability = DC_PAPERS
    '�o�b�t�@�̕K�v�ȃT�C�Y���擾
    lngSupportedPapersNeeded = DeviceCapabilities(strPrinterDeviceName, _
                                                  strPrinterPortName, _
                                                  lngDeviceCapability, _
                                                  ByVal vbNullString, _
                                                  ByVal vbNullString)
    '�o�b�t�@���m��
    ReDim intSupportedPapers(lngSupportedPapersNeeded - 1)
    '�T�|�[�g����Ă���p���T�C�Y���擾
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, _
                                               strPrinterPortName, _
                                               lngDeviceCapability, _
                                               intSupportedPapers(0), _
                                               ByVal vbNullString)
    '�o�b�t�@�ɗp���T�C�Y�����݂����
    For lngSupportedPapersCount = 0 To lngSupportedPapersNeeded - 1
        If (cboPaperSize.ListCount - 1) < lngSupportedPapersCount Then Exit For
        If cboPaperSize.Text = cboPaperSize.List(lngSupportedPapersCount) Then
            GetPaperSize = intSupportedPapers(lngSupportedPapersCount)
            Exit For
        End If
    Next lngSupportedPapersCount

    Exit Function
    
GetPaperSize_Err:

    GetPaperSize = 0
    Call MsgBox("�p���T�C�Y�R�[�h�擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "GetPaperSize_Err")
    
End Function

Private Function ListPaperSizes() As Boolean
    
    Dim strPrinterDeviceName         As String
    Dim strPrinterPortName           As String
    Dim lngDeviceCapability          As Long
    Dim lngSupportedPaperNamesNeeded As Long
    Dim bytSupportedPaperNames()     As Byte
    Dim strSupportedPaperName        As String * 64
    Dim lngSupportedPaperNamesCount  As Long
    Dim lngWin32apiResultCode        As Long
    
    On Error GoTo ListPaperSizes_Err
    
    ListPaperSizes = False
    
    '�R���{�{�b�N�X��������
    cboPaperSize.Clear
    
    If optDeviceName1.Value = True Then
        '�ʏ�g���v�����^�̏ꍇ
        '�f�o�C�X���ƃ|�[�g�����w��
        With Printer
            strPrinterDeviceName = .DeviceName
            strPrinterPortName = .Port
        End With
    ElseIf optDeviceName2.Value = True Then
        If Trim(cboDeviceName.Text) = "" Then Exit Function
        '�w��̃v�����^�̏ꍇ
        Dim objPrinter As Printer
        
        For Each objPrinter In Printers
            If Trim(objPrinter.DeviceName) = Trim(cboDeviceName.Text) Then
                strPrinterDeviceName = objPrinter.DeviceName
                strPrinterPortName = objPrinter.Port
                Exit For
            End If
        Next
    End If
    
    '�₢���킹��\�͂��w��
    lngDeviceCapability = DC_PAPERNAMES
    '�o�b�t�@�ɕK�v�ȃT�C�Y���擾
    lngSupportedPaperNamesNeeded = DeviceCapabilities(strPrinterDeviceName, strPrinterPortName, lngDeviceCapability, ByVal vbNullString, ByVal vbNullString)
    '�o�b�t�@���m��
    ReDim bytSupportedPaperNames(64 - 1, lngSupportedPaperNamesNeeded - 1)
    '�g�p�ł���p���̖��O���擾
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, strPrinterPortName, lngDeviceCapability, bytSupportedPaperNames(0, 0), ByVal vbNullString)
    '�g�p�ł���p���̖��O���
    For lngSupportedPaperNamesCount = 0 To lngSupportedPaperNamesNeeded - 1
        '�p���̖��O�𕡎�
        MoveMemory ByVal strSupportedPaperName, bytSupportedPaperNames(0, lngSupportedPaperNamesCount), 64
        '�p���̖��O��\��
        cboPaperSize.AddItem Left(strSupportedPaperName, InStr(strSupportedPaperName, vbNullChar) - 1)
    Next lngSupportedPaperNamesCount

    ListPaperSizes = True

    Exit Function

ListPaperSizes_Err:

    ListPaperSizes = False
    Call MsgBox("�p���T�C�Y�ꗗ�擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListPaperSizes_Err")
    
End Function

Private Function GetPaperSource(ByVal intPaperSource As Integer) As Boolean
    
    Dim intIndex1 As Integer
    
    On Error GoTo GetPaperSource_Err
    
    GetPaperSource = False
    
    If ListPaperSource() = False Then Exit Function
    
    With cboPaperSource
        .Text = .List(0)    '�������@�u�����v�������\��
        
        For intIndex1 = 0 To .ListCount - 1
            If .ItemData(intIndex1) = intPaperSource Then
                .Text = .List(intIndex1)
            End If
        Next intIndex1
    End With
            
    GetPaperSource = True

    Exit Function

GetPaperSource_Err:

    GetPaperSource = False
    Call MsgBox("�������@���擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "GetPaperSource_Err")
    
End Function

Private Function ListPaperSource() As Boolean
    
    Dim strPrinterDeviceName       As String
    Dim strPrinterPortName         As String
    Dim lngDeviceCapability        As Long
    Dim lngSupportedBinNamesNeeded As Long
    Dim bytSupportedBinNames()     As Byte
    Dim strSupportedBinName        As String * 24
    Dim lngSupportedBinNamesCount  As Long
    Dim lngWin32apiResultCode      As Long
    Dim intSupportedBins()         As Integer
    
    On Error GoTo ListPaperSource_Err
    
    ListPaperSource = False
    
    '�R���{�{�b�N�X��������
    cboPaperSource.Clear
    
    '�f�o�C�X���ƃ|�[�g�����w��
    If optDeviceName1.Value = True Then
        With Printer
            strPrinterDeviceName = .DeviceName
            strPrinterPortName = .Port
        End With
    Else
        If Trim$(cboDeviceName.Text) = "" Then Exit Function
        strPrinterDeviceName = cboDeviceName.Text
'        strPrinterPortName = lblPortName.Caption
    End If
    
    '�₢���킹��\�͂��w��
    lngDeviceCapability = DC_BINNAMES
    
    '�o�b�t�@�ɕK�v�ȃT�C�Y���擾
    lngSupportedBinNamesNeeded = DeviceCapabilities(strPrinterDeviceName, _
                                                    strPrinterPortName, _
                                                    lngDeviceCapability, _
                                                    ByVal vbNullString, _
                                                    ByVal vbNullString)
    '�o�b�t�@���m��
    ReDim bytSupportedBinNames(24 - 1, lngSupportedBinNamesNeeded - 1)
    
    '�g�p�ł��鋋�����@�̖��O���擾
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, _
                                               strPrinterPortName, _
                                               lngDeviceCapability, _
                                               bytSupportedBinNames(0, 0), _
                                               ByVal vbNullString)
    
    '�₢���킹��\�͂��w��
    lngDeviceCapability = DC_BINS
    
    '�o�b�t�@�ɕK�v�ȃT�C�Y���擾
    lngSupportedBinNamesNeeded = DeviceCapabilities(strPrinterDeviceName, _
                                                    strPrinterPortName, _
                                                    lngDeviceCapability, _
                                                    ByVal vbNullString, _
                                                    ByVal vbNullString)
    ' �o�b�t�@���m��
    ReDim intSupportedBins(lngSupportedBinNamesNeeded - 1)
    
    '�g�p�ł��鋋�����@���擾
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, _
                                               strPrinterPortName, _
                                               lngDeviceCapability, _
                                               intSupportedBins(0), _
                                               ByVal vbNullString)
    
    '�g�p�ł��鋋�����@�̖��O���
    With cboPaperSource
        '�o�b�t�@�ɋ������@�̖��O�������
        For lngSupportedBinNamesCount = 0 To lngSupportedBinNamesNeeded - 1
            '�������@�̖��O��؂�o��
            MoveMemory ByVal strSupportedBinName, bytSupportedBinNames _
                       (0, _
                       lngSupportedBinNamesCount), _
                        24
            
            '�������@�̖��O��\��
            .AddItem Left(strSupportedBinName, _
                     InStr(strSupportedBinName, _
                     vbNullChar) - 1)
            
            '�������@�R�[�h�̐ݒ�
            .ItemData(lngSupportedBinNamesCount) = intSupportedBins(lngSupportedBinNamesCount)
        
        Next lngSupportedBinNamesCount
    End With
    
    ListPaperSource = True

    Exit Function

ListPaperSource_Err:

    ListPaperSource = False
    Call MsgBox("�������@�ꗗ�擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListPaperSource_Err")
    
End Function

Private Function GetCboListName(ByRef objCboControl As Control, ByVal varData As Variant) As Boolean
    
    Dim intIndex1 As Integer
    
    On Error GoTo GetCboListName_Err
    
    GetCboListName = False
    
    With objCboControl
        For intIndex1 = 0 To .ListCount - 1
            If CStr(.List(intIndex1)) = CStr(varData) Then
                .Text = .List(intIndex1)
                GetCboListName = True
                Exit Function
            End If
        Next intIndex1
    End With

    Exit Function

GetCboListName_Err:

    GetCboListName = False
    Call MsgBox("�R���{�{�b�N�X���疼�̎擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "GetCboListName_Err")
    
End Function
