VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptMT070_2 
   ClientHeight    =   11010
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptMT070_2.dsx":0000
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptMT070_2.dsx":000C
End
Attribute VB_Name = "rptMT070_2"
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
    frmPrintDialog.objArPrint.NoData = True
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
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.RecordSet
    
    On Error GoTo ActiveReport_ReportStart_Err
        
    With adoRecordset1
        '�ݒ�}�X�^
        strSQL = "{call sp_MT010;1}"
        .Open strSQL, frmMt070.m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            txtCaddress.Text = IIf(IsNull(.Fields("CAddres")), "", Trim(.Fields("CAddres")))
            txtCname.Text = IIf(IsNull(.Fields("Company")), "", Trim(.Fields("Company")))
            txtCceo.Text = IIf(IsNull(.Fields("CCeo")), "", Trim(.Fields("CCeo")))
            txtCTel.Text = "TEL " & IIf(IsNull(.Fields("CTel")), "", Trim(.Fields("CTel")))
            txtCfax.Text = "FAX " & IIf(IsNull(.Fields("CFax")), "", Trim(.Fields("CFax")))
            txtCurl.Text = "URL " & IIf(IsNull(.Fields("CUrl")), "", Trim(.Fields("CUrl")))
            txtCyubin.Text = IIf(IsNull(.Fields("CPost")), "", Trim(.Fields("CPost")))
        Else
            txtCaddress.Text = ""
            txtCname.Text = ""
            txtCceo.Text = ""
            txtCTel.Text = ""
            txtCfax.Text = ""
            txtCurl.Text = ""
            txtCyubin.Text = ""
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    '�X�֔ԍ�
    txtCyubin1.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 1, 1)
    txtCyubin2.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 2, 1)
    txtCyubin3.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 3, 1)
    txtCyubin4.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 4, 1)
    txtCyubin5.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 5, 1)
    txtCyubin6.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 6, 1)
    txtCyubin7.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 7, 1)
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("���|�[�g�������J�n���钼�O�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub Detail_Format()

    On Error GoTo Detail_Format_Err
    
    '�X�֔ԍ�
    txtYubin1.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 1, 1)
    txtYubin2.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 2, 1)
    txtYubin3.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 3, 1)
    txtYubin4.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 4, 1)
    txtYubin5.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 5, 1)
    txtYubin6.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 6, 1)
    txtYubin7.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 7, 1)
    
    '����
    If Trim(txtCeo.Text) <> "" And Trim(txtCeo.Text) <> Trim(txtBname.Text) Then
        txtBnameStr.Text = Trim(txtBname.Text) & vbCrLf & Trim(txtCeo.Text) & "�@�l"
    Else
        txtBnameStr.Text = Trim(txtBname.Text) & "�@�l"
    End If
    
    '����R�[�h
    txtBcodeStr.Text = "(" & txtBcode.Text & ")"
    
    Exit Sub
    
Detail_Format_Err:
    
   Call MsgBox("�ڍ׃t�H�[�}�b�g�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Format_Err")

End Sub
