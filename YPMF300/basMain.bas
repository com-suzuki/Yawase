Attribute VB_Name = "basMain"
Option Explicit

Public g_clsAdoSQL As New clsAdoCore
Public g_clsAdoAccess As New clsAdoCore
Public g_clsReg As New clsReg
Public g_blnLoginOK As Boolean          'ログインフラグ
Public g_strPcode As String             '担当者コード
Public g_strPname As String             '担当者名
Public g_strOdate As String             '開催年月日

Sub Main()
    
    On Error GoTo Main_Err
    
    '重複起動のチェック
    If App.PrevInstance = True Then
        End
    End If

    'レジストリ読み込み
    g_clsReg.RegKey = REG_KEY
    If g_clsReg.ReadReg = False Then
        End
    End If

    'データベース接続
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
    
    Call MsgBox("プログラム実行エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "Main_Err")
    
End Sub
