VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�ް�ޮݏ��"
   ClientHeight    =   2895
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6210
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1998.18
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   5831.511
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   60
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   337.12
      ScaleMode       =   0  'հ�ް
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   0
      Top             =   2400
      Width           =   5880
   End
   Begin VB.Label lblUrl 
      Alignment       =   2  '��������
      Caption         =   "http://www.com-e.co.jp"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1320
      MouseIcon       =   "frmAbout.frx":0D06
      MousePointer    =   99  'հ�ް��`
      TabIndex        =   5
      Top             =   1920
      Width           =   3345
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "������Ё@�R���E�G���W�j�A�����O"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      TabIndex        =   4
      Top             =   1500
      Width           =   5955
   End
   Begin VB.Label lblTitle 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   5325
   End
   Begin VB.Label lblVersion 
      Caption         =   "Ver 1.00"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   3
      Top             =   900
      Width           =   5295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�w�肳�ꂽ�t�@�C�����I�[�v���A���邢�͕\������API
Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Private Const SW_SHOW = 5

'�f�X�N�g�b�v�E�B���h�E�̃n���h�����擾����API
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Sub cmdOK_Click()
  
  Unload Me

End Sub

Private Sub Form_Load()

    lblTitle.Caption = SYSTEM_NAME
    lblVersion.Caption = "Ver " & PROGRAM_VERSION

End Sub

Private Sub lblUrl_Click()

    Dim lngAPIReVal As Long
    Dim strFileName As String

    On Error GoTo lblUrl_Click_Err
    
    'URL���J��
    strFileName = Trim(lblUrl.Caption)
    lngAPIReVal = ShellExecute(GetDesktopWindow, "open", strFileName, Chr$(0), "", SW_SHOW)

    Exit Sub

lblUrl_Click_Err:

    Call MsgBox("���x���N���b�N���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "lblUrl_Click_Err")

End Sub
