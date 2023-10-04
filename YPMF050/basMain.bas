Attribute VB_Name = "basMain"
Option Explicit

Public g_clsAdoSQL As New clsAdoCore
Public g_clsAdoAccess As New clsAdoCore
Public g_clsReg As New clsReg
Public g_blnLoginOK As Boolean          'ログインフラグ
Public g_strPcode As String             '担当者コード
Public g_strPname As String             '担当者名
Public g_strOdate As String             '開催年月日

Public g_strBcode As String             '買主コード
Public g_strRePrintNum As String        '回数

Sub Main()
    
    On Error GoTo Main_Err
    
    '重複起動のチェック
    If Command() = "" Then
        If App.PrevInstance = True Then
            End
        End If
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

    'コマンドライン引数
    '開催日
    '担当者コード
    '担当者名
    '買主コード
    '再発行番号
    If Command() = "" Then
        g_blnLoginOK = False
        g_strPcode = ""
        g_strPname = ""
        g_strOdate = ""
        g_strBcode = ""
        g_strRePrintNum = "0"
        
        frmLogin.Show vbModal
        If g_blnLoginOK = False Then End
        frmYpmf050.Show
        Unload frmLogin
    Else
        Dim varCommnad() As String
        varCommnad = Split(Command(), ",")
        
        g_blnLoginOK = True
        g_strPcode = varCommnad(1)
        g_strPname = varCommnad(2)
        g_strOdate = varCommnad(0)
        g_strBcode = varCommnad(3)
        g_strRePrintNum = varCommnad(4)

        frmYpmf050.Show
    End If
    
    Exit Sub
    
Main_Err:
    
    Call MsgBox("プログラム実行エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "Main_Err")
    
End Sub
