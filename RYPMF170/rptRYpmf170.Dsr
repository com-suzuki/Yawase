VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptRYpmf170 
   ClientHeight    =   11010
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptRYpmf170.dsx":0000
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptRYpmf170.dsx":000C
End
Attribute VB_Name = "rptRYpmf170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnMeisaiPrint As Boolean

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
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    
    On Error GoTo ActiveReport_ReportStart_Err
       
    '������t
    Me.PrintDay.Text = Format(Now(), "yyyy/mm/dd")
    With adoRecordset1
        '�ݒ�}�X�^
        strSQL = "{call sp_MT010;1}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            txtCname.Text = IIf(IsNull(.Fields("Company")), "", Trim(.Fields("Company")))
        Else
            txtCname.Text = ""
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Detail.Visible = frmRYpmf170.optPrint(0).Value
    txtOdate.Visible = frmRYpmf170.optPrint(1).Value
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("���|�[�g�������J�n���钼�O�G���[�I�I" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub PageHeader_Format()
    
    txtBnameStr.Text = Trim(txtBname.Text) & "�@�l"
    
End Sub
