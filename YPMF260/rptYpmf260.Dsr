VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf260 
   ClientHeight    =   11115
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf260.dsx":0000
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptYpmf260.dsx":000C
End
Attribute VB_Name = "rptYpmf260"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
       
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("���|�[�g�������J�n���钼�O�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub Detail_Format()

    On Error GoTo Detail_Format_Err
       
    If frmYpmf260.optDiv(1).Value = True Then
        txtUriage_Kingaku.Visible = False
        LineUriage_Kingaku2.Visible = False
    Else
        txtUriage_Kingaku.Visible = True
        LineUriage_Kingaku2.Visible = True
    End If
       
    Exit Sub
    
Detail_Format_Err:
    
   Call MsgBox("���׃G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Format_Err")

End Sub

Private Sub GroupFooter1_Format()

    On Error GoTo GroupFooter1_Format_Err
       
    '���v�^�C�g�����Ȃ��ꍇ�͕\�����Ȃ�
    If Trim(txtSubtotal_Name.Text) = "" Then
        GroupFooter1.Visible = False
    Else
        GroupFooter1.Visible = True
    End If
       
    If frmYpmf260.optDiv(1).Value = True Then
        txtUriage_Kingaku_Subtotal.Visible = False
        LineUriage_Kingaku3.Visible = False
    Else
        txtUriage_Kingaku_Subtotal.Visible = True
        LineUriage_Kingaku3.Visible = True
    End If
       
    Exit Sub
    
GroupFooter1_Format_Err:
    
   Call MsgBox("�O���[�v�t�b�^�[�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupFooter1_Format_Err")
  
End Sub

Private Sub PageHeader_Format()

    On Error GoTo PageHeader_Format_Err
       
    If frmYpmf260.optDiv(1).Value = True Then
        lblUriage_Kingaku.Visible = False
        LineUriage_Kingaku1.Visible = False
    Else
        lblUriage_Kingaku.Visible = True
        LineUriage_Kingaku1.Visible = True
    End If
       
    Exit Sub
    
PageHeader_Format_Err:
    
   Call MsgBox("�y�[�W�w�b�_�[�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageHeader_Format_Err")
  
End Sub

Private Sub ReportFooter_Format()

    On Error GoTo ReportFooter_Format_Err
       
    txtTokki.Text = frmYpmf260.txtTokki.Text
       
    If frmYpmf260.optDiv(1).Value = True Then
        txtUriage_Kingaku_Total.Visible = False
        LineUriage_Kingaku4.Visible = False
    Else
        txtUriage_Kingaku_Total.Visible = True
        LineUriage_Kingaku4.Visible = True
    End If
       
    Exit Sub
    
ReportFooter_Format_Err:
    
   Call MsgBox("���|�[�g�t�b�^�[�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ReportFooter_Format_Err")
  
End Sub
