VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf060_2 
   ClientHeight    =   11010
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf060_2.dsx":0000
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptYpmf060_2.dsx":000C
End
Attribute VB_Name = "rptYpmf060_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngDataCount As Long    '��������̃f�[�^��

'�ځ@�I�@�@�F
'���@���@�@�F���|�[�g�G���[��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F
'�X�V�����@�F
'
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports.IReturnBool)

    '�G���[��\������
    CancelDisplay = False

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�f�[�^���Ȃ��ꍇ
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F
'�X�V�����@�F
'
Private Sub ActiveReport_NoData()

    On Error Resume Next

    Call MsgBox("�f�[�^������܂���B", vbOKOnly + vbInformation, "���")
    Me.Cancel

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F���|�[�g�������J�n���钼�O
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F
'�X�V�����@�F
'
Private Sub ActiveReport_ReportStart()
    
    On Error GoTo ActiveReport_ReportStart_Err
       
    '������t
    Me.PrintDay.Text = Format(Now(), "yyyy/mm/dd")
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("���|�[�g�������J�n���钼�O�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub Detail_Format()
    
    lngDataCount = lngDataCount + 1

End Sub

Private Sub GroupFooter3_Format()
    
    If lngDataCount > 1 Then
        GroupFooter3.Visible = True
    Else
        GroupFooter3.Visible = False
    End If
    lngDataCount = 0
    
End Sub

Private Sub PageHeader_Format()
    
    On Error GoTo PageHeader_Format_Err
       
    txtFrom.Text = frmPrintDialog.m_strRdateFrom
    txtTo.Text = frmPrintDialog.m_strRdateTo
       
    Exit Sub
    
PageHeader_Format_Err:
    
   Call MsgBox("�y�[�W�w�b�_�[�t�H�[�}�b�g���G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageHeader_Format_Err")

End Sub
