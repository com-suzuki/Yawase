VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf120 
   ClientHeight    =   11115
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf120.dsx":0000
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptYpmf120.dsx":000C
End
Attribute VB_Name = "rptYpmf120"
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
       
    '������t
    Me.PrintDay.Text = Format(Now(), "yyyy/mm/dd")
    '���y�[�W����
    GroupFooter1.NewPage = ddNPAfter
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("���|�[�g�������J�n���钼�O�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub GroupFooter1_Format()
    
    On Error GoTo GroupFooter1_Format_Err
       
    '���y�[�W����
    If Me.DataControl1.Recordset.EOF = True Then
        GroupFooter1.NewPage = ddNPNone
    End If
    
    If Trim(txtGTotal1.Text) <> "" And Trim(txtNyukin1.Text) <> "" Then
       txtZandaka1.Text = CCur(txtGTotal1.Text) - CCur(txtNyukin1.Text)
       txtZandaka1.Text = Format(txtZandaka1.Text, "#,##0")
    End If
    
    Exit Sub
    
GroupFooter1_Format_Err:
    
   Call MsgBox("�O���[�v�t�b�^�[�t�H�[�}�b�g���G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupFooter1_Format_Err")
  
End Sub

Private Sub GroupFooter2_Format()
    
    On Error GoTo GroupFooter2_Format_Err
       
    If Trim(txtGtotal.Text) <> "" And Trim(txtNyukin.Text) <> "" Then
       txtZandaka.Text = CCur(txtGtotal.Text) - CCur(txtNyukin.Text)
       txtZandaka.Text = Format(txtZandaka.Text, "#,##0")
    End If
       
    Exit Sub
    
GroupFooter2_Format_Err:
    
   Call MsgBox("�O���[�v�t�b�^�[�t�H�[�}�b�g���G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupFooter2_Format_Err")
       
End Sub

Private Sub PageHeader_Format()
     
    On Error GoTo PageHeader_Format_Err
       
    If txtDiv.Text = "0" Then
        lblTitle.Visible = True
        txtOdateTitle.Visible = True
        lineTitle.Visible = True
    Else
        lblTitle.Visible = False
        txtOdateTitle.Visible = False
        lineTitle.Visible = False
    End If
           
    Exit Sub
    
PageHeader_Format_Err:
    
   Call MsgBox("�y�[�W�w�b�_�[�t�H�[�}�b�g���G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageHeader_Format_Err")
  

End Sub

Private Sub ReportFooter_Format()
        
    On Error GoTo ReportFooter_Format_Err
       
    If Trim(txtGtotal_Total.Text) <> "" And Trim(txtNyukin_total.Text) <> "" Then
       txtZandaka_Total.Text = CCur(txtGtotal_Total.Text) - CCur(txtNyukin_total.Text)
       txtZandaka_Total.Text = Format(txtZandaka_Total.Text, "#,##0")
    End If
       
    Exit Sub
    
ReportFooter_Format_Err:
    
   Call MsgBox("���|�[�g�t�b�^�[�t�H�[�}�b�g���G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ReportFooter_Format_Err")
    
End Sub
