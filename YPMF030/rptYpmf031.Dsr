VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf031 
   ClientHeight    =   11115
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf031.dsx":0000
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptYpmf031.dsx":000C
End
Attribute VB_Name = "rptYpmf031"
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

    Me.GroupHeader2.NewPage = ddNPNone

    Exit Sub

ActiveReport_ReportStart_Err:

   Call MsgBox("���|�[�g�������J�n���钼�O�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")

End Sub

Private Sub Detail_Format()
    
    On Error GoTo Detail_Format_Err

    txtMsg.Visible = False
    txtPrice.Visible = True
    If Trim(txtIdiv.Text) <> "" And Trim(txtIdiv.Text) <> "0" And Trim(txtIdiv.Text) <> Trim(txtLine.Text) Then
        txtPrice.Visible = False
        txtMsg.Visible = True
        txtMsg.Text = "(���F" & txtIdiv.Text & "�ɍ��Z)"
    End If
    If Trim(txtResult.Text) <> "" And Trim(txtResult.Text) <> "0" Then
        txtPrice.Visible = False
        txtMsg.Visible = True
        txtMsg.Text = "(�����s����)"
    End If
    
    Exit Sub

Detail_Format_Err:

   Call MsgBox("���׃t�H�[�}�b�g�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Format_Err")

End Sub

Private Sub GroupFooter1_Format()

    Dim curBuff As Currency

    On Error GoTo GroupFooter1_Format_Err
   
    '202308
'    txtSasihiki.Text = "0"
'    If IsNumeric(txtTotal.Text) = True And IsNumeric(txtCharge.Text) = True Then
'        curBuff = CCur(txtTotal.Text) - CCur(txtCharge.Text)
'        txtSasihiki.Text = Format(curBuff, "#,##0")
'    End If
    '202308

    If txtChargeDisplay.Text = "��0" Then
        txtChargeDisplay.Text = "0"
    End If
    If txtKeep.Text = "��0" Then
        txtKeep.Text = "0"
    End If
    
    '202308 �C���{�C�X�Ή��ύX
    lblTotal.Caption = Format(CDbl(txtSubTotal.DataValue) + CDbl(txtChargeDisplay.DataValue) + CDbl(txtKeep.DataValue), "###,###,###,##0")

    Exit Sub

GroupFooter1_Format_Err:

   Call MsgBox("�O���[�v�t�b�^�[�t�H�[�}�b�g�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupFooter1_Format_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�O���[�v�t�b�^�[�����
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F
'�X�V�����@�F
'
Private Sub GroupHeader2_AfterPrint()

    On Error GoTo GroupHeader2_AfterPrint_Err

    Me.GroupHeader2.NewPage = ddNPBefore

    Exit Sub

GroupHeader2_AfterPrint_Err:

   Call MsgBox("�O���[�v�t�b�^�[�����G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupHeader2_AfterPrint_Err")

End Sub

'�ځ@�I�@�@�F
'���@���@�@�F�y�[�W�t�b�^�[�t�H�[�}�b�g��
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F
'�X�V�����@�F
'
Private Sub PageFooter_Format()

    On Error GoTo PageFooter_Format_Err

    Exit Sub

PageFooter_Format_Err:

   Call MsgBox("�y�[�W�t�b�^�[�t�H�[�}�b�g���G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageFooter_Format_Err")

End Sub

Private Sub PageHeader_BeforePrint()
    
    On Error GoTo PageHeader_BeforePrint_Err
    
    txtPnum.Text = txtPnumFooter.Text
    txtSname.Text = txtSnameFooter.Text
    '202308
    txtTnum.Text = txtTnumFooter.Text

    Exit Sub

PageHeader_BeforePrint_Err:

   Call MsgBox("�y�[�W�w�b�_�[����O�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageHeader_BeforePrint_Err")

End Sub
