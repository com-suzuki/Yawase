Attribute VB_Name = "basMain"
Option Explicit

Public g_clsReg As New clsReg

Sub Main()

    'スプラッシュ画面表示
    frmSplash.Show vbModeless
    frmSplash.Refresh
    
    'レジストリ読み込み
    g_clsReg.RegKey = REG_KEY
    If g_clsReg.ReadReg = False Then
        Call MsgBox("環境が設定されていません。", vbOKOnly + vbCritical, "")
        Call Shell(App.Path & "\" & "Config.exe", vbNormalFocus)
    End If
    
    JustWait    '待機
    Unload frmSplash
    
    frmMenu.Show vbModeless

End Sub

Private Sub JustWait()

    Dim StartTime, Inval
    
    StartTime = Now()
    'インターバル　２秒
    Inval = #12:00:02 AM#
    Do While (Now < StartTime + Inval)
        DoEvents
    Loop
    
End Sub
