Attribute VB_Name = "basMain"
Option Explicit

Public g_clsReg As New clsReg

Sub Main()

    '�X�v���b�V����ʕ\��
    frmSplash.Show vbModeless
    frmSplash.Refresh
    
    '���W�X�g���ǂݍ���
    g_clsReg.RegKey = REG_KEY
    If g_clsReg.ReadReg = False Then
        Call MsgBox("�����ݒ肳��Ă��܂���B", vbOKOnly + vbCritical, "")
        Call Shell(App.Path & "\" & "Config.exe", vbNormalFocus)
    End If
    
    JustWait    '�ҋ@
    Unload frmSplash
    
    frmMenu.Show vbModeless

End Sub

Private Sub JustWait()

    Dim StartTime, Inval
    
    StartTime = Now()
    '�C���^�[�o���@�Q�b
    Inval = #12:00:02 AM#
    Do While (Now < StartTime + Inval)
        DoEvents
    Loop
    
End Sub
