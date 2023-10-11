VERSION 5.00
Object = "{E95678BE-E45E-471F-9287-59E8911E479E}#1.5#0"; "ArViewer15j.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmARPreview 
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   Icon            =   "frmARPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   9285
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   420
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer arv 
      Height          =   7665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   13520
      SectionData     =   "frmARPreview.frx":0CFA
   End
End
Attribute VB_Name = "frmARPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_objArPrint As clsArPrint

Const TOOLBAR_ADD01 = "�y�[�W�ݒ�"
Const TOOLBAR_ADD02 = "PDF�o��"
Const TOOLBAR_ADD03 = "����"

Private Sub arv_PrintAborted()

    Unload Me
    
End Sub

Private Sub arv_ToolbarClick(ByVal tool As DDActiveReportsViewerCtl.DDTool)

    Select Case tool.Caption
        Case TOOLBAR_ADD01: '�y�[�W�ݒ�
            Set frmARPage.m_objArPrint = m_objArPrint
            frmARPage.Show vbModal
        Case TOOLBAR_ADD02: 'PDF�o��
            Call OutPutPDF
        Case TOOLBAR_ADD03: '�t�H�[�������
            Unload Me
        Case Else
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    With arv.ToolBar
        '�c�[���o�[�̈ꕔ�̃A�C�e�����\���ɂ��܂��B
        .Tools(0).Visible = False    '���o��
        .Tools(12).Visible = False   '�߂�
        .Tools(13).Visible = False   '�i��
    
        '�c�[���o�[�̒ǉ�
        .Tools.Add TOOLBAR_ADD01
        .Tools(.Tools.Count - 1).Type = 0
        .Tools.Add ""
        .Tools(.Tools.Count - 1).Type = 2
        .Tools.Add TOOLBAR_ADD02
        .Tools(.Tools.Count - 1).Type = 0
        .Tools.Add ""
        .Tools(.Tools.Count - 1).Type = 2
        .Tools.Add TOOLBAR_ADD03
        .Tools(.Tools.Count - 1).Type = 0
        .Tools.Add ""
        .Tools(.Tools.Count - 1).Type = 2
    End With
    
End Sub

Private Sub Form_Resize()
           
    On Error Resume Next
           
    If Me.WindowState <> vbMinimized Then
        '�t�H�[���̃T�C�Y�ύX�ɂ��킹�ĤARViewer�R���g���[�������T�C�Y���܂��B
        With arv
            .Top = 0
            .Left = 0
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight
        End With
    End If

End Sub

Private Function OutPutPDF()
           
    Dim strFileName As String
           
    On Error GoTo OutPutPDF_Err
           
    '�o�͐�w��
    CommonDialog1.FileName = m_objArPrint.Caption & ".pdf"
    CommonDialog1.DefaultExt = "pdf"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "pdf�t�@�C�� (*.pdf)|*.pdf|���ׂẴt�@�C�� (*.*)|*.*"
    On Error Resume Next
    CommonDialog1.ShowSave
    If Err.Number = cdlCancel Then
        Exit Function
    End If
    On Error GoTo OutPutPDF_Err
    strFileName = CommonDialog1.FileName
    If strFileName = "" Then Exit Function
    
    Screen.MousePointer = vbHourglass

    Dim objARExport As New ARExportPDF
    Dim objARExportFix As New ARExportPDFFix
    
    objARExport.FileName = strFileName
    m_objArPrint.objReport.Run
    m_objArPrint.objReport.Export objARExport
    
    'PDF�t�@�C���̕����������C��
    objARExportFix.ExportFix objARExport.FileName

    Screen.MousePointer = vbDefault

    Exit Function

OutPutPDF_Err:

    Call MsgBox("PDF�o�̓G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "OutPutPDF_Err")

End Function

