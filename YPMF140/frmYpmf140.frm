VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmYpmf140 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   3105
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf140.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   12150
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdLogin 
         Caption         =   "�J�ÔN�����ƒS���҂̕ύX"
         Height          =   375
         Left            =   6960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9900
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   675
         _Version        =   262145
         _ExtentX        =   1191
         _ExtentY        =   635
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "2;20"
         Contents        =   "frmYpmf140.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�J�ÔN����"
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   8
         Left            =   3480
         TabIndex        =   15
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�S����"
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
      Begin VB.Label lblPname 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "�m�m�m�m�m�m�m�m�m�m"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   4980
         TabIndex        =   13
         Top             =   180
         Width           =   1905
      End
      Begin VB.Label lblOdate 
         Alignment       =   2  '��������
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "9999/12/31"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1620
         TabIndex        =   12
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   9
      Top             =   2280
      Width           =   12015
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "��ʸر(F8)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   2
         rPic.top        =   4
         rPic.right      =   0
         rPic.bottom     =   0
         rText.left      =   10
         rText.top       =   8
         rText.right     =   103
         rText.bottom    =   27
         Picture         =   "frmYpmf140.frx":0D13
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10200
         TabIndex        =   5
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "�I��(F9)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
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
         Picture         =   "frmYpmf140.frx":0D2F
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   8460
         TabIndex        =   4
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "���(F12)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
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
         Picture         =   "frmYpmf140.frx":0E89
      End
   End
   Begin VB.Frame fraMeisai 
      Height          =   1635
      Left            =   60
      TabIndex        =   8
      Top             =   660
      Width           =   12015
      Begin CSComboLib.CSComboBox cboBcode 
         Height          =   405
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Top             =   180
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf140.frx":0F9B
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "����R�[�h"
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
         LabelWidth      =   67
         LabelHeight     =   25
         LabelLeft       =   15
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
      Begin CSComboLib.CSComboBox cboBcode 
         Height          =   405
         Index           =   1
         Left            =   1620
         TabIndex        =   2
         Top             =   1020
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf140.frx":0FB4
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��������
         Caption         =   "�`"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   20
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblBcode_Name 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   2700
         TabIndex        =   19
         Top             =   1020
         Width           =   9195
      End
      Begin VB.Label lblBcode_Name 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����
         Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   2700
         TabIndex        =   18
         Top             =   180
         Width           =   9195
      End
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   12240
      TabIndex        =   0
      Top             =   120
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf140.frx":0FCD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf140.frx":103B
      Key             =   "frmYpmf140.frx":1059
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
      Left            =   12240
      TabIndex        =   6
      Top             =   240
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf140.frx":109D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf140.frx":110B
      Key             =   "frmYpmf140.frx":1129
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
      TabIndex        =   7
      Top             =   1200
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   132
      Caption         =   "frmYpmf140.frx":116D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf140.frx":11DB
      Key             =   "frmYpmf140.frx":11F9
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
Attribute VB_Name = "frmYpmf140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_strGinkoName  As String

Private Sub cboBcode_Click(Index As Integer)

    Call cboBcode_Validate(Index, False)
    
End Sub

Private Sub cboBcode_DropDown(Index As Integer)

    Call MakecboBcode(cboBcode(Index))
    
End Sub

Private Sub cboBcode_GotFocus(Index As Integer)

    cboBcode(Index).BackColor = FOCUS_STOP_COLOR
    cboBcode(Index).Tag = cboBcode(Index).Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboBcode_LostFocus(Index As Integer)
   
    cboBcode(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboBcode_Validate(Index As Integer, Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cboBcode_Validate_Err
    
    If Trim(cboBcode(Index).Text) = "" Then Exit Sub
    If cboBcode(Index).Tag = cboBcode(Index).Text Then Exit Sub
    
    lblBcode_Name(Index).Caption = ""
    
    With adoRecordset1
        '���Ӑ�}�X�^
        strSQL = "{call sp_MT070;2(" & Trim(cboBcode(Index).Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Fdiv")) Then
                If .Fields("Fdiv") = BUSINESS_DIV_BUYER Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    lblBcode_Name(Index).Caption = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If Index = 0 Then
        cboBcode(1).Text = cboBcode(0).Text
        lblBcode_Name(1).Caption = lblBcode_Name(0).Caption
    End If
    
    Exit Sub

cboBcode_Validate_Err:

    Call MsgBox("�t�H�[�J�X�ړ��O�G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBcode_Validate_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F��ʃN���A�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    cboBcode(0).SetFocus

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���s�N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If MsgBox("���s���܂����H", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    '���̓`�F�b�N
    If DoValidationChecks() = False Then Exit Sub
    '����p���[�N�쐬
    If MakePrintWork() = False Then Exit Sub
    '����v���r���[
    If ActiveReportPrint(0) = False Then Exit Sub

    Exit Sub
    
cmdExecute_Click_Err:

    Call MsgBox("���s�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�I���N���b�N��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdLogin_Click()

    On Error GoTo cmdLogin_Click_Err

    g_blnLoginOK = False
    g_strPcode = Trim(cboPcode.Text)
    g_strPname = lblPname.Caption
    g_strOdate = Trim(lblOdate.Caption)
    frmLogin.Show vbModal
    If g_blnLoginOK = True Then
        lblOdate.Caption = g_strOdate
        cboPcode.Text = g_strPcode
        lblPname.Caption = g_strPname
    End If
    Unload frmLogin
    
    Exit Sub

cmdLogin_Click_Err:

    Call MsgBox("�J�ÔN�����ƒS���҂̕ύX�N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdLogin_Click_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[���L�[�_�E����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Form_KeyDown_Err
    
    '���^�[���L�[�Ŏ��̃R���g���[���փt�H�[�J�X�ړ�
    If KeyCode = vbKeyReturn And Shift <> vbShiftMask Then
        KeyCode = 0
        Global_SendKeys Me, VK_TAB
        Exit Sub
    End If
    
    '�V���[�g�J�b�g�L�[�̊��蓖��
    Select Case KeyCode
        Case vbKeyF1
        Case vbKeyF2
        Case vbKeyF3
        Case vbKeyF4
        Case vbKeyF5
        Case vbKeyF6
        Case vbKeyF7
        Case vbKeyF8
            cmdClear.SetFocus
            DoEvents
            Call cmdClear_Click
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

    Call MsgBox("�t�H�[���L�[�_�E�����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[�����[�h��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    Me.Caption = SYSTEM_NAME & "-" & "�������������o��"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname
    
    Call FieldsClear(0)
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("�t�H�[�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    End

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�t�H�[���A�����[�h��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set g_clsAdoSQL = Nothing
    Set g_clsReg = Nothing
    End
    
    Exit Sub
    
Form_Unload_Err:

    Call MsgBox("�t�H�[���A�����[�h���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")
    End

End Sub

'�ځ@�I�@�@�F��ʃN���A
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F0�F�S���
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        cboBcode(0).Text = "0"
        cboBcode(0).Tag = ""
        cboBcode(1).Text = "9999"
        cboBcode(1).Tag = ""
        lblBcode_Name(0).Caption = ""
        lblBcode_Name(1).Caption = ""
'        Call FieldsSet
    End If
    
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("��ʃN���A�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

'�ځ@�I�@�@�F���̓`�F�b�N
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "�J�ÔN��������͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(cboBcode(0).Text) = "" Then
        cboBcode(0).SetFocus
        strErrMsg = "����R�[�h����͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    If Trim(cboBcode(1).Text) = "" Then
        cboBcode(1).SetFocus
        strErrMsg = "����R�[�h����͂��Ă��������B"
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "���̓`�F�b�N")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("���̓`�F�b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function

'�ځ@�I�@�@�F����p���[�N�쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim adoDT041 As New ADODB.Recordset
    Dim adoDT041_M As New ADODB.Recordset
    Dim adoDT060 As New ADODB.Recordset
    Dim adoDT021 As New ADODB.Recordset
    Dim adoMT010 As New ADODB.Recordset
    Dim adoMT070 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim strBuff As String
    Dim intIndex1 As Integer

    Dim intPage As Integer
    Dim intLine As Integer

    Dim strKey As String
    Dim strBcode As String
    Dim strBname As String
    Dim strPost As String
    Dim strAdress As String
    Dim curSeikyu_Total As Currency
    Dim curPrice_Total As Currency
    Dim curNyukin_Total As Currency
    Dim strKessai_Date As String
    
    Const PAGE_MAX_LINE = 24                    '1�y�[�W�̍ő�s��

    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    '��Аݒ�}�X�^
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        m_strGinkoName = RTrim(adoMT010("Memo"))
    Else
        m_strGinkoName = ""
    End If
    adoMT010.Close
    
    '���[�N�폜
    strSQL = "DELETE FROM WK_YPMF140"
    g_clsAdoAccess.Connection.Execute strSQL

    '���[�N�I�[�v��
    strSQL = "SELECT * FROM WK_YPMF140"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    '���吸�Z�f�[�^�I�[�v��
    strSQL = "{call sp_YPMF1401;1('" & Trim(lblOdate.Caption) & "'," & cboBcode(0).Text & "," & cboBcode(1).Text & ")}"
    adoDT041.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoDT041.EOF = True Then
        Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If
        
    frmCount.fpProgressBar1.Value = 0
    frmCount.fpProgressBar1.Max = adoDT041.RecordCount
    frmCount.Show
    Me.Enabled = False
    
    Do While Not adoDT041.EOF
        strBcode = Format$(adoDT041.Fields("Bcode"), "0000")
        strBname = Global_Get_Bname(g_clsAdoSQL, strBcode, Trim(lblOdate.Caption), strBuff) & "�@�l"
        
        '���Ӑ�}�X�^
        strSQL = "{call sp_MT070;2(" & strBuff & ")}"
        adoMT070.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If adoMT070.EOF = False Then
            strPost = IIf(IsNull(adoMT070.Fields("Post")), "", Trim(adoMT070.Fields("Post")))
            strAdress = IIf(IsNull(adoMT070.Fields("Addres")), "", Trim(adoMT070.Fields("Addres")))
        Else
            strPost = ""
            strAdress = ""
        End If
        adoMT070.Close
        
        '������
        curSeikyu_Total = 0
        '���������v
        curPrice_Total = IIf(IsNull(adoDT041.Fields("Gtotal")), 0, adoDT041.Fields("Gtotal"))
        curNyukin_Total = 0
        strKessai_Date = Format$(Now(), "yyyy�Nmm��dd��") & "���ϕ�"
        
        intPage = 1
        intLine = 1
        
        '���吸�Z�f�[�^�I�[�v��
        strSQL = "{call sp_YPMF1402;1('" & Trim(lblOdate.Caption) & "'," & adoDT041.Fields("Bcode") & ")}"
        adoDT041_M.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT041_M.EOF
            
            '�w�b�_�[
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "�J�Ó��F" & Format$(adoDT041_M.Fields("Odate"), "yyyy�Nmm��dd��")
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = Null
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            '�y�[�W�v�Z
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
'********** ������ **********
            
            '�������׃f�[�^�I�[�v��
            strSQL = "{call sp_YPMF1403;1('" & Format$(adoDT041_M.Fields("Odate"), "yyyymmdd") & "'," & adoDT041_M.Fields("Bcode") & "," & adoDT041_M.Fields("Num") & ")}"
            adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT021.EOF Then
                adoDT021.Close
                
                '�f�[�^���Ȃ��ꍇ�͋������׃f�[�^�ݐς�T��
                strSQL = "{call sp_YPMF1403;2('" & Format$(adoDT041_M.Fields("Odate"), "yyyymmdd") & "'," & adoDT041_M.Fields("Bcode") & "," & adoDT041_M.Fields("Num") & ")}"
                adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            End If
            
            Do While Not adoDT021.EOF
                wkRecordset.AddNew
                wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
                wkRecordset.Fields("PageNum") = intPage
                wkRecordset.Fields("Line") = intLine
                
                wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
                wkRecordset.Fields("Bname") = strBname
                wkRecordset.Fields("Post") = strPost
                wkRecordset.Fields("Addres") = strAdress
                wkRecordset.Fields("Price_Total") = 0
                wkRecordset.Fields("Nyukin_Total") = 0
                wkRecordset.Fields("Kessai_Date") = strKessai_Date
                
                wkRecordset.Fields("Iname") = adoDT021.Fields("Iname")
                wkRecordset.Fields("Qty") = adoDT021.Fields("Qty")
                wkRecordset.Fields("Price") = adoDT021.Fields("Price")
                If adoDT021.Fields("Qty") <> 0 Then
                    wkRecordset.Fields("Tanka") = Fix(CCur(wkRecordset.Fields("Price")) / CCur(wkRecordset.Fields("Qty")))
                Else
                    wkRecordset.Fields("Tanka") = 0
                End If
                wkRecordset.Fields("Tekiyo") = ""
                wkRecordset.Update
                
                '�y�[�W�v�Z
                intLine = intLine + 1
                If intLine > PAGE_MAX_LINE Then
                    intPage = intPage + 1
                    intLine = 1
                End If
                
                adoDT021.MoveNext
            Loop
            adoDT021.Close
                    
'********** ������ **********
            
             '�������׃f�[�^
            strSQL = "SELECT * FROM DT031" & _
                     " WHERE Odate = '" & adoDT041_M.Fields("Odate") & "'" & _
                     " AND Bcode = " & adoDT041_M.Fields("Bcode") & _
                     " AND Bnum = " & adoDT041_M.Fields("Num") & _
                     " ORDER BY Onum, Line"
            adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If adoDT021.EOF Then
                adoDT021.Close
            
                '�f�[�^���Ȃ��ꍇ�͒������׃f�[�^�ݐς�T��
                strSQL = "SELECT * FROM RT031" & _
                         " WHERE Odate = '" & adoDT041_M.Fields("Odate") & "'" & _
                         " AND Bcode = " & adoDT041_M.Fields("Bcode") & _
                         " AND Bnum = " & adoDT041_M.Fields("Num") & _
                         " ORDER BY Onum, Line"
                adoDT021.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            End If
            
            Do While Not adoDT021.EOF
                wkRecordset.AddNew
                wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
                wkRecordset.Fields("PageNum") = intPage
                wkRecordset.Fields("Line") = intLine
                
                wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
                wkRecordset.Fields("Bname") = strBname
                wkRecordset.Fields("Post") = strPost
                wkRecordset.Fields("Addres") = strAdress
                wkRecordset.Fields("Price_Total") = 0
                wkRecordset.Fields("Nyukin_Total") = 0
                wkRecordset.Fields("Kessai_Date") = strKessai_Date
                
                wkRecordset.Fields("Iname") = adoDT021.Fields("Iname")
                wkRecordset.Fields("Qty") = adoDT021.Fields("Qty")
                wkRecordset.Fields("Price") = adoDT021.Fields("Price")
                If adoDT021.Fields("Qty") <> 0 Then
                    wkRecordset.Fields("Tanka") = Fix(CCur(wkRecordset.Fields("Price")) / CCur(wkRecordset.Fields("Qty")))
                Else
                    wkRecordset.Fields("Tanka") = 0
                End If
                wkRecordset.Fields("Tekiyo") = ""
                wkRecordset.Update
                
                '�y�[�W�v�Z
                intLine = intLine + 1
                If intLine > PAGE_MAX_LINE Then
                    intPage = intPage + 1
                    intLine = 1
                End If
                
                adoDT021.MoveNext
            Loop
            adoDT021.Close
                    
            '���v
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "           ���@�v"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Total")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            '�y�[�W�v�Z
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
            '�����
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "           �����"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Tax")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            '�y�[�W�v�Z
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
            '�ێ��Ǘ���
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "           �ێ��Ǘ���"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Keep")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            '�y�[�W�v�Z
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
            '201107 �����萔��
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "           �����萔��"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Brate2")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            '�y�[�W�v�Z
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
            
            '���v
            wkRecordset.AddNew
            wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
            wkRecordset.Fields("PageNum") = intPage
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
            wkRecordset.Fields("Bname") = strBname
            wkRecordset.Fields("Post") = strPost
            wkRecordset.Fields("Addres") = strAdress
            wkRecordset.Fields("Price_Total") = 0
            wkRecordset.Fields("Nyukin_Total") = 0
            wkRecordset.Fields("Kessai_Date") = strKessai_Date
            wkRecordset.Fields("Iname") = "********** ���@�v **********"
            wkRecordset.Fields("Qty") = Null
            wkRecordset.Fields("Price") = adoDT041_M.Fields("Gtotal")
            wkRecordset.Fields("Tanka") = Null
            wkRecordset.Fields("Tekiyo") = ""
            wkRecordset.Update
            
            '�y�[�W�v�Z
            intLine = intLine + 1
            If intLine > PAGE_MAX_LINE Then
                intPage = intPage + 1
                intLine = 1
            End If
        
            adoDT041_M.MoveNext
        Loop
        adoDT041_M.Close
        
        '���吸�Z�f�[�^�I�[�v��
        strSQL = "SELECT Odate FROM DT041" & _
                 " WHERE Odate <= '" & Trim(lblOdate.Caption) & "'" & _
                 " AND Bcode = " & adoDT041.Fields("Bcode") & _
                 " AND Rdiv = 0 " & _
                 " GROUP BY Odate" & _
                 " ORDER BY Odate"
        adoDT041_M.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT041_M.EOF
            '�����f�[�^�I�[�v��
            strSQL = "{call sp_YPMF1404;1('" & adoDT041_M.Fields("Odate") & "'," & adoDT041.Fields("Bcode") & ")}"
            adoDT060.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            Do While Not adoDT060.EOF
                '����
                wkRecordset.AddNew
                wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
                wkRecordset.Fields("PageNum") = intPage
                wkRecordset.Fields("Line") = intLine
                wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
                wkRecordset.Fields("Bname") = strBname
                wkRecordset.Fields("Post") = strPost
                wkRecordset.Fields("Addres") = strAdress
                wkRecordset.Fields("Price_Total") = 0
                wkRecordset.Fields("Nyukin_Total") = 0
                wkRecordset.Fields("Kessai_Date") = strKessai_Date
                
                Select Case adoDT060.Fields("R")
                    Case "1":
                        wkRecordset.Fields("Iname") = Format(adoDT060.Fields("Rdate"), "yyyy�Nmm��dd��") & "�@����(����)"
                    Case "2":
                        wkRecordset.Fields("Iname") = Format(adoDT060.Fields("Rdate"), "yyyy�Nmm��dd��") & "�@����(���؎�)"
                    Case "3":
                        wkRecordset.Fields("Iname") = Format(adoDT060.Fields("Rdate"), "yyyy�Nmm��dd��") & "�@����(��s�U��)"
                End Select
                
                wkRecordset.Fields("Qty") = Null
                wkRecordset.Fields("Price") = adoDT060.Fields("Ptotal")
                wkRecordset.Fields("Tanka") = Null
                wkRecordset.Fields("Tekiyo") = ""
                wkRecordset.Update
                
                '�y�[�W�v�Z
                intLine = intLine + 1
                If intLine > PAGE_MAX_LINE Then
                    intPage = intPage + 1
                    intLine = 1
                End If
            
                '�������z�v�Z
                curNyukin_Total = curNyukin_Total + CCur(adoDT060.Fields("Ptotal"))
                    
                adoDT060.MoveNext
            Loop
            adoDT060.Close
            
            adoDT041_M.MoveNext
        Loop
        adoDT041_M.Close
            
        '��s�쐬(�ŏI�s�܂Ńf�[�^������ꍇ��intLine���P�ƂȂ�)
        If intLine <> 1 Then
            For intIndex1 = intLine To PAGE_MAX_LINE
                wkRecordset.AddNew
                wkRecordset.Fields("Key") = strBcode & "-" & Format$(intPage, "000")
                wkRecordset.Fields("PageNum") = intPage
                wkRecordset.Fields("Line") = intIndex1
                wkRecordset.Fields("Bcode") = "(" & strBcode & ")"
                wkRecordset.Fields("Bname") = strBname
                wkRecordset.Fields("Post") = strPost
                wkRecordset.Fields("Addres") = strAdress
                wkRecordset.Fields("Price_Total") = 0
                wkRecordset.Fields("Nyukin_Total") = 0
                wkRecordset.Fields("Kessai_Date") = strKessai_Date
                wkRecordset.Fields("Iname") = Null
                wkRecordset.Fields("Qty") = Null
                wkRecordset.Fields("Price") = Null
                wkRecordset.Fields("Tanka") = Null
                wkRecordset.Fields("Tekiyo") = Null
                wkRecordset.Update
            Next intIndex1
        End If
       
        '�������z���v�Z�i���������v�|�������v�j
        curSeikyu_Total = curPrice_Total - curNyukin_Total

        If curSeikyu_Total > 0 Then
            '�����z�X�V
            strSQL = "UPDATE WK_YPMF140"
            strSQL = strSQL & " SET Seikyu_Total = " & curSeikyu_Total & ","
            strSQL = strSQL & " Price_Total = " & curPrice_Total & ","
            strSQL = strSQL & " Nyukin_Total = " & curNyukin_Total
            strSQL = strSQL & " WHERE Bcode = '" & "(" & strBcode & ")" & "'"
            g_clsAdoAccess.Connection.Execute strSQL
        Else
            '�c�����Ȃ��ꍇ�̓��[�N�f�[�^�폜
            strSQL = "DELETE FROM WK_YPMF140"
            strSQL = strSQL & " WHERE Bcode = '" & "(" & strBcode & ")" & "'"
            g_clsAdoAccess.Connection.Execute strSQL
        End If

        adoDT041.MoveNext
        
        DoEvents
        frmCount.fpProgressBar1.Value = frmCount.fpProgressBar1.Value + 1
        If frmCount.g_blnCancel Then GoTo MakePrintWork_Cancel:
    Loop
    adoDT041.Close
    
    wkRecordset.Close
    
    '�o�O�h�~
    strSQL = "SELECT * FROM WK_YPMF140"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    wkRecordset.Requery
    
    If wkRecordset.EOF = True Then
        wkRecordset.Close
        Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If
    
    wkRecordset.Close
    
    MakePrintWork = True
    
MakePrintWork_Exit:
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Unload frmCount
    
    Exit Function

MakePrintWork_Cancel:

    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    MakePrintWork = False
    Call MsgBox("������[�N�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function

'�ځ@�I�@�@�F�R���{�{�b�N�X�̍쐬
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Sub MakecboBcode(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboBcode_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        strSQL = "SELECT * FROM vw_MT071" & _
                 " WHERE (Fdiv = " & BUSINESS_DIV_BUYER & " OR Fdiv = " & BUSINESS_DIV_ALL & ")"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            If Not IsNull(.Fields("Sdate")) And Not IsNull(.Fields("Fdate")) Then
                If .Fields("Sdate") <= Trim(lblOdate.Caption) And Trim(lblOdate.Caption) <= .Fields("Fdate") Then
                    Ctrl.AddItem .Fields("Bnum") & ";" & .Fields("Bname")
                Else
                    Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
                End If
            Else
                Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
            End If
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboBcode_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("�R���{�{�b�N�X�쐬�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboBcode_Err")

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    cboBcode(0).SetFocus

End Sub

'�ځ@�I�@�@�FActiveReport�̈��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F0:�v���r���[ 1:���
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf140
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "������������"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .Caption = "������������"
        If .PrintActiveReport(intFlg) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End With

    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    ActiveReportPrint = True
    
    Exit Function
    
ActiveReportPrint_Err:

    ActiveReportPrint = False
    Screen.MousePointer = vbDefault
    Call MsgBox("���s�N���b�N�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReportPrint_Err")
    
End Function

'�ځ@�I�@�@�F�t�B�[���h�̃Z�b�g
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�R�^�O�U�^�P�W
'�X�V�����@�F
'
Private Function FieldsSet() As Boolean
    
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim strBuff As String

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    '���吸�Z�f�[�^
    strSQL = "SELECT * FROM MT070" & _
             " WHERE Fdiv = " & BUSINESS_DIV_BUYER & " OR Fdiv = " & BUSINESS_DIV_ALL & _
             " ORDER BY Bcode"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        cboBcode(0).Text = adoRecordset1.Fields("Bcode")
        lblBcode_Name(0).Caption = Global_Get_Bname(g_clsAdoSQL, adoRecordset1.Fields("Bcode"), lblOdate.Caption, strBuff)
        adoRecordset1.MoveLast
        cboBcode(1).Text = adoRecordset1.Fields("Bcode")
        lblBcode_Name(1).Caption = Global_Get_Bname(g_clsAdoSQL, adoRecordset1.Fields("Bcode"), lblOdate.Caption, strBuff)
    End If
    
    adoRecordset1.Close
    Set adoRecordset1 = Nothing
    
    Screen.MousePointer = vbDefault
    
    FieldsSet = True
    
    Exit Function

FieldsSet_Err:

    FieldsSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("�t�B�[���h�Z�b�g�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsSet_Err")

End Function

