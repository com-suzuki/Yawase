Attribute VB_Name = "basMain"
Option Explicit

Public g_clsAdoSQL As New clsAdoCore
Public g_clsAdoAccess As New clsAdoCore
Public g_clsReg As New clsReg
Public g_blnLoginOK As Boolean          '���O�C���t���O
Public g_strPcode As String             '�S���҃R�[�h
Public g_strPname As String             '�S���Җ�
Public g_strOdate As String             '�J�ÔN����

Sub Main()
    
    On Error GoTo Main_Err
    
    '�d���N���̃`�F�b�N
    If App.PrevInstance = True Then
        End
    End If

    '���W�X�g���ǂݍ���
    g_clsReg.RegKey = REG_KEY
    If g_clsReg.ReadReg = False Then
        End
    End If

    '�f�[�^�x�[�X�ڑ�
    With g_clsAdoSQL
        .Provider = adoSQLServer
        .Server = g_clsReg.Server
        .DBName = g_clsReg.DBName
        .UID = g_clsReg.UID
        .PWD = g_clsReg.PWD
        .CommandTimeOut = g_clsReg.CommandTimeOut
        If .Connect = False Then
            End
        End If
    End With
    With g_clsAdoAccess
        .Provider = adoAccess
        .DBName = g_clsReg.LDatabase & "\" & g_clsReg.LDBName
        If .Connect = False Then
            End
        End If
    End With

    g_blnLoginOK = False
    g_strPcode = ""
    g_strPname = ""
    g_strOdate = ""
    frmLogin.Show vbModal
    If g_blnLoginOK = False Then End
    frmYpmf300.Show
    Unload frmLogin
    
    Exit Sub
    
Main_Err:
    
    Call MsgBox("�v���O�������s�G���[�I�I" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "Main_Err")
    
End Sub
