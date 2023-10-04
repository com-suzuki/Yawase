Attribute VB_Name = "basGlobal"
Option Explicit

'スリープ関数
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetVersion Lib "kernel32" () As Long

'APIでNumLockをOnにする
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'APIでキー入力を行う
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const VK_NUMLOCK = &H90
Public Const VK_CAPSLOCK = &H14
Public Const KEYEVENTF_KEYUP = &H2
Public Const WM_KEYDOWN = &H100     'キーダウン
Public Const VK_TAB = &H9           'TAB
Public Const VK_RETURN = &HD        'Enter
Public Const VK_SHIFT = &H10        'Shift

'プログラム終了まで待機
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As Long, ByVal lpCommandLine As Long, ByVal IDProcess As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal lpdExitCode As Long, hHandle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ACTIVE = &H103

'SHBrowseForFolderqで使用する構造体
Private Type BROWSEINFO
    hwndOwner As Long       '親Windowのﾊﾝﾄﾞﾙ
    pidlRoot As Long        'ﾙｰﾄﾌｫﾙﾀﾞ
    pszDisplayName As Long
    lpszTitle As String     'ﾀﾞｲｱﾛｸﾞに表示するﾒｯｾｰｼﾞ
    ulFlags As Long         'ｵﾌﾟｼｮﾝ
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'ﾙｰﾄﾌｫﾙﾀﾞ定数
Private Const CSIDL_DESKTOP = &H0           'ﾃﾞｽｸﾄｯﾌﾟ
Private Const CSIDL_PROGRAMS = &H2          'ﾌﾟﾛｸﾞﾗﾑ
Private Const CSIDL_CONTROLS = &H3          'ｺﾝﾄﾛｰﾙﾊﾟﾈﾙ
Private Const CSIDL_PRINTERS = &H4          'ﾌﾟﾘﾝﾀｰ
Private Const CSIDL_PERSONAL = &H5          'ﾊﾟｰｿﾅﾙ
Private Const CSIDL_FAVORITES = &H6         'ﾌﾞｯｸﾏｰｸ
Private Const CSIDL_STARTUP = &H7           'ｽﾀｰﾄｱｯﾌﾟ
Private Const CSIDL_RECENT = &H8            '[最近使ったﾌｧｲﾙ]
Private Const CSIDL_SENDTO = &H9            '[送る]
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB         '[ｽﾀｰﾄ]ﾒﾆｭｰ
Private Const CSIDL_DESKTOPDIRECTORY = &H10 'ﾃﾞｽｸﾄｯﾌﾟ
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOO = &H13           'Network Neighborhood
Private Const CSIDL_FONTS = &H14            'ﾌｫﾝﾄ
Private Const CSIDL_TEMPLATES = &H15        'Shell New

'特殊ﾌｫﾙﾀﾞ(ﾏｲｺﾝﾋﾟｭｰﾀ、ｺﾝﾄﾛｰﾙﾊﾟﾈﾙ等)を選択させない
Private Const BIF_BROWSEFORCOMPUTER = 1
'[ﾌｫﾙﾀﾞの参照]ﾀﾞｲｱﾛｸﾞを呼び出すAPI
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFO As BROWSEINFO) As Long
'SHBrowseForFolderで得られた値からﾌｫﾙﾀﾞのﾊﾟｽを取得するAPI
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'SHBrowseForFolderで得られた値のﾒﾓﾘを開放するAPI
Private Declare Function SHFree Lib "shell32" Alias "#195" (ByVal pidl As Long) As Long
'テンポラリファイルのために指定されているパスを取得する関数の宣言
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'クラス名とウィンドウ名が指定された文字列と一致するウィンドウのハンドルを取得する関数の宣言
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'ウィンドウがアイコン化されているかどうか判断する関数の宣言
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
'ウィンドウを復元してアクティブ化する関数の宣言
Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
'指定したウィンドウを一番手前に持ってくる関数の宣言
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

'INIファイル操作関数宣言
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type RECT
    left As Long    'WindowのX座標
    top As Long     'WindowのY座標
    right As Long   'Windowの右端の座標
    bottom As Long  'Windowの底にあたる部分の座標
End Type
Private Const HWND_TOP = 0           '手前にｾｯﾄ
Private Const HWND_BOTTOM = 1        '後ろにｾｯﾄ
Private Const HWND_TOPMOST = -1      '常に手前にｾｯﾄ
Private Const HWND_NOTOPMOST = -2    '常に手前、解除
Private Const SWP_SHOWWINDOW = &H40  '表示する

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'目　的　　：Windowsのバージョン判定
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function IsWindowsNT() As Boolean

    IsWindowsNT = IIf(GetVersion() And &H80000000, False, True)

End Function

'目　的　　：SendKeys（SendKey("{TAB}")を２回連続して実行するとNumLockがはずれるバグ回避）
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Sub Global_SendKeys(ByVal OFORM As Object, ByVal CHARLNG As Long)

    Dim varState As Variant

    'NumLock情報取得
    varState = GetKeyState(VK_NUMLOCK)

    'Next Fieldへ
    Call PostMessage(OFORM.hwnd, WM_KEYDOWN, CHARLNG, 0)
    
    'NumLock On
    If varState <> 0 And GetKeyState(VK_NUMLOCK) = 0 Then
        Call keybd_event(VK_NUMLOCK, 0, 0, 0)
        Call keybd_event(VK_NUMLOCK, 0, KEYEVENTF_KEYUP, 0)
    End If

End Sub

'目　的　　：小数点以下四捨五入
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_Round(ByVal dblArg1 As Double) As Double

    If dblArg1 >= 0 Then
        dblArg1 = CCur(dblArg1) + 0.5
    Else
        dblArg1 = CCur(dblArg1) - 0.5
    End If
    Global_Round = Fix(CCur(dblArg1))

End Function

'目　的　　：小数点以下切り上げ
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_RoundUp(ByVal dblArg1 As Double) As Double

    If dblArg1 >= 0 Then
        dblArg1 = CCur(dblArg1) + 0.9999
    Else
        dblArg1 = CCur(dblArg1) - 0.9999
    End If
    Global_RoundUp = Fix(CCur(dblArg1))

End Function

'目　的　　：外部アプリケーションが終了するまで待機する
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_Exec_Application(ByRef strCmdline As String, ByRef intWindowstyle As Integer) As Boolean
    
    Dim hShell As Long
    Dim hProc As Long
    Dim lExit As Long
    Dim bret As Long

    On Error Resume Next

    hShell = Shell(strCmdline, intWindowstyle)
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hShell)
    Do
        GetExitCodeProcess hProc, lExit
        DoEvents
    Loop While lExit = STILL_ACTIVE
    
    bret = CloseHandle(hProc)

    Global_Exec_Application = True

    Exit Function
 
End Function

'目　的　　：月の取得
'条　件　　：
'結　果　　：
'引　数　　：lngMonth：対象となる年月　dblDiff：何ヶ月後（前）
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_Get_PassingMonth(ByVal lngMonth As Long, ByVal intDiff As Integer) As Long

    Dim intYear As Integer
    Dim intMonth As Integer
    Dim varDate As Variant
    
    On Error Resume Next
    
    '年と月に分解
    intYear = left(lngMonth, 4)
    intMonth = right(lngMonth, 2)
    '日付を求める
    varDate = DateSerial(intYear, intMonth + intDiff, 1)
    
    Global_Get_PassingMonth = CLng(Format(varDate, "yyyymm"))
  
End Function

'目　的　　：日の取得
'条　件　　：
'結　果　　：
'引　数　　：lngMonth：対象となる年月日　dblDiff：何日後（前）
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_Get_PassingDay(ByVal lngYyyymmdd As Long, ByVal intDiff As Integer) As Long

    Dim intYear As Integer
    Dim intMonth As Integer
    Dim intDay As Integer
    Dim varDate As Variant
    
    On Error Resume Next
    
    '年月日に分解
    intYear = left(lngYyyymmdd, 4)
    intMonth = Mid(lngYyyymmdd, 5, 2)
    intDay = right(lngYyyymmdd, 2)
    
    '日付を求める
    varDate = DateSerial(intYear, intMonth, intDay + intDiff)
    
    Global_Get_PassingDay = CLng(Format(varDate, "yyyymmdd"))
  
End Function

'目　的　　：月の最終日取得
'条　件　　：
'結　果　　：
'引　数　　：lngMonth：対象となる年月
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_Get_MonthLastDay(ByVal lngMonth As Long) As Integer

    Dim intYear As Integer
    Dim intMonth As Integer
    Dim varDate As Variant
    
    On Error Resume Next
    
    '年と月に分解
    intYear = left(lngMonth, 4)
    intMonth = right(lngMonth, 2)
    '日付を求める
    varDate = DateSerial(intYear, intMonth + 1, 1 - 1)
    
    Global_Get_MonthLastDay = Day(varDate)
    
End Function

'目　的　　：Unicode→Ansi
'条　件　　：
'結　果　　：
'引　数　　：strArg：文字列　intByte：バイト数
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_LeftB_Ansi(ByRef strArg As String, ByRef intByte As Integer) As String

    On Error Resume Next

    If strArg = "" Or intByte = 0 Then
        Global_LeftB_Ansi = ""
    End If

    Global_LeftB_Ansi = StrConv(LeftB(StrConv(strArg, vbFromUnicode), intByte), vbUnicode)

End Function

'目　的　　：Unicode→Ansi
'条　件　　：
'結　果　　：
'引　数　　：strArg：文字列　intByte：バイト数
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_RightB_Ansi(ByRef strArg As String, ByRef intByte As Integer) As String

    On Error Resume Next

    If strArg = "" Or intByte = 0 Then
        Global_RightB_Ansi = ""
    End If

    Global_RightB_Ansi = StrConv(RightB(StrConv(strArg, vbFromUnicode), intByte), vbUnicode)

End Function

'目　的　　：Unicode→Ansi
'条　件　　：
'結　果　　：
'引　数　　：strArg：文字列　intByte：バイト数
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_MidB_Ansi(ByRef strArg As String, ByRef intByte1 As Integer, ByRef intByte2 As Integer) As String

    On Error Resume Next

    If strArg = "" Or intByte1 = 0 Or intByte2 = 0 Then
        Global_MidB_Ansi = ""
    End If

    Global_MidB_Ansi = StrConv(MidB(StrConv(strArg, vbFromUnicode), intByte1, intByte2), vbUnicode)

End Function

'目　的　　：Unicode→Ansi
'条　件　　：
'結　果　　：
'引　数　　：strArg：文字列　intByte：バイト数
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_LenB_Ansi(ByRef strArg As String) As Integer

    On Error Resume Next

    If strArg = "" Then
        Global_LenB_Ansi = 0
    End If

    Global_LenB_Ansi = LenB(StrConv(strArg, vbFromUnicode))

End Function

'目　的　　：月差の取得
'条　件　　：
'結　果　　：
'引　数　　：lngMonth1：対象となる年月１ lngMonth2：対象となる年月２
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Global_Get_DiffMonth(ByVal lngMonth1 As Long, ByVal lngMonth2 As Long) As Integer

    Dim strDate1 As String
    Dim strDate2 As String
    
    On Error Resume Next
    
    strDate1 = left(Format(lngMonth1, "000000"), 4) & "/" & right(Format(lngMonth1, "000000"), 2) & "/01"
    strDate2 = left(Format(lngMonth2, "000000"), 4) & "/" & right(Format(lngMonth2, "000000"), 2) & "/01"
    
    Global_Get_DiffMonth = Abs(DateDiff("m", Format(CDate(strDate1), "yyyy/mm/dd"), Format(CDate(strDate2), "yyyy/mm/dd")))
  
End Function

'目　的　　：フォルダ選択ダイアログ表示
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function OpenSelectFolderDialog(ByRef hwnd As Long) As String
    
    Dim typBrowseInfo As BROWSEINFO
    Dim lngFoldPointer As Long
    Dim strPathName As String

    On Error GoTo OpenSelectFolderDialog_Err

    OpenSelectFolderDialog = ""

    With typBrowseInfo
        '親ウインドウを設定
        .hwndOwner = hwnd
        'ルートフォルダを設定
        .pidlRoot = 0
        .lpszTitle = "フォルダ選択"
        '特殊フォルダを選択させない
        .ulFlags = BIF_BROWSEFORCOMPUTER
    End With

    '[フォルダの参照]ダイアログを呼び出す
    lngFoldPointer = SHBrowseForFolder(typBrowseInfo)
    If lngFoldPointer = 0 Then Exit Function

    '予めNull文字をセット
    strPathName = String$(256, vbNullChar)
    'SHBrowseForFolderで得られた値からフォルダのパスを取得
    Call SHGetPathFromIDList(lngFoldPointer, strPathName)

    '割り当てられたメモリを開放
    Call SHFree(lngFoldPointer)

    If Trim(strPathName) <> "" Then
        OpenSelectFolderDialog = strPathName
    End If
    
    Exit Function
    
OpenSelectFolderDialog_Err:

    OpenSelectFolderDialog = ""
    Call MsgBox("フォルダ選択ダイアログエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "OpenSelectFolderDialog_Err")

End Function

'目　的　　：yyyy/mm/dd(文字列)をyyyymmdd(数値)に変換する
'条　件　　：不正日付チェックなし
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／０５
'更新履歴　：
'
Public Function Global_Get_NumericDay(ByRef strYyyymmdd As String) As Long

    Dim varBuff As Variant
    Dim strDay As String
    Dim intIndex1 As Integer
    
    On Error Resume Next
    
    varBuff = Split(strYyyymmdd, "/")
    strDay = ""
    For intIndex1 = 0 To 2
        strDay = strDay & varBuff(intIndex1)
    Next intIndex1
    
    Global_Get_NumericDay = CLng(strDay)
  
End Function

'目　的　　：yyyymmdd(数値)をyyyy/mm/dd(文字列)に変換する
'条　件　　：不正日付チェックなし
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／０５
'更新履歴　：
'
Public Function Global_Get_StringDay(ByRef lngYyyymmdd As Long) As String

    Dim strDay As String
    
    On Error Resume Next
    
    strDay = left$(Format(lngYyyymmdd, "00000000"), 4) & "/"
    strDay = strDay & Mid$(Format(lngYyyymmdd, "00000000"), 5, 2) & "/"
    strDay = strDay & right$(Format(lngYyyymmdd, "00000000"), 2)
    
    Global_Get_StringDay = strDay
  
End Function

'目　的　　：yyyy mm ddをyyyy/mm/ddにする
'条　件　　：不正日付チェックなし
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／０５
'更新履歴　：
'
Public Function Global_StrToDate(ByRef strYyyy As String, ByRef strMm As String, ByRef strDd As String) As String

    On Error Resume Next
    Global_StrToDate = Trim(strYyyy) & "/" & Format(strMm, "00") & "/" & Format(strDd, "00")
  
End Function

'目　的　　：重複起動のチェック
'条　件　　：
'結　果　　：
'引　数　　：クラス名／タイトル／実行形式ファイルパス／実行ファイル
'戻り値　　：正常＝True／エラー＝False
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００１／０３／０６
'更新履歴　：
'
Public Function Global_IsLoad(strClassName As String, strAppTitle As String, strBinPath As String, strExe As String) As Boolean

    Dim lHandleWindow As Long
    Dim varResponse  As Variant

    Global_IsLoad = True

    On Error GoTo Global_IsLoad_Err

    lHandleWindow = FindWindow(strClassName, strAppTitle)
    If lHandleWindow <> 0 Then
        If IsIconic(lHandleWindow) <> False Then
            varResponse = OpenIcon(lHandleWindow)
        Else
            varResponse = BringWindowToTop(lHandleWindow)
        End If
    Else
        If strBinPath <> "" And strExe <> "" Then
            varResponse = Shell(strBinPath & strExe, vbNormalFocus)
        End If
    End If

    Exit Function

Global_IsLoad_Err:
    
    Global_IsLoad = False

End Function

'目　的　　：INIファイルからのデータ取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００１／０３／０５
'更新履歴　：
'
Public Function Global_GetIni(ByVal vstrAppname As String, ByVal vstrKeyword As String, ByVal vstrIniFile As String) As String
    
    Dim strResult   As String * 1024
    Dim intTemp     As Integer
    Dim intLen      As Integer
    Dim strFileName As String
   
    On Error Resume Next
   
    intTemp = GetPrivateProfileString(vstrAppname, vstrKeyword, "", strResult, Len(strResult), vstrIniFile)
    Global_GetIni = left$(strResult, intTemp)

End Function

'目　的　　：INIファイルへデータ更新
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００１／０３／０５
'更新履歴　：
'
Public Function Global_SetIni(ByVal vstrAppname As String, ByVal vstrKeyword As String, ByVal vstrKeyVal As String, ByVal vstrIniFile As String) As Integer
    
    Dim intTemp As Integer

    On Error Resume Next

    intTemp = WritePrivateProfileString(vstrAppname, vstrKeyword, vstrKeyVal, vstrIniFile)
    Global_SetIni = intTemp

End Function

'目　的　　：日付のチェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／０５
'更新履歴　：
'
Public Function Global_IsDate(ByRef strYyyy As String, ByRef strMm As String, ByRef strDd As String) As Boolean

    On Error Resume Next
    Global_IsDate = IsDate(Global_StrToDate(strYyyy, strMm, strDd))
  
End Function

'目　的　　：フォームを常に手前に表示を設定／解除する
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／０５
'更新履歴　：
'
Public Function Global_SetFormTop(lngHWnd As Long, bolIsTop As Boolean)
    
    Dim mLeft As Long
    Dim mTop As Long
    Dim mWidth As Long
    Dim mHeight As Long
    Dim lpRect As RECT

    On Error GoTo Global_SetFormTop_Err

    Global_SetFormTop = GetWindowRect(lngHWnd, lpRect)

    mLeft = lpRect.left
    mTop = lpRect.top
    mWidth = lpRect.right - lpRect.left
    mHeight = lpRect.bottom - lpRect.top
    
    If bolIsTop Then
        Global_SetFormTop = SetWindowPos(lngHWnd, HWND_TOPMOST, mLeft, mTop, mWidth, mHeight, SWP_SHOWWINDOW)
    Else
        Global_SetFormTop = SetWindowPos(lngHWnd, HWND_NOTOPMOST, mLeft, mTop, mWidth, mHeight, SWP_SHOWWINDOW)
    End If

    Exit Function

Global_SetFormTop_Err:

    Call MsgBox("フォームを常に手前に表示を設定／解除エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_SetFormTop_Err")

End Function

Public Function Global_StrNull(ByVal varString As Variant, Optional ByVal strDefault As String = "") As String

    On Error Resume Next

    Global_StrNull = IIf(IsNull(varString) = True, strDefault, varString)

End Function

Public Function Global_GetPcName() As String

    Dim strBuff As String
    Dim intPos As Integer
    
    Const MAX_COMPUTERNAME_LENGTH = 255

    On Error Resume Next

    strBuff = Space(MAX_COMPUTERNAME_LENGTH + 1)
    GetComputerName strBuff, MAX_COMPUTERNAME_LENGTH

    'Chr$(0)以降の文字を削除
    intPos = InStr(strBuff, Chr$(0))
    If intPos > 0 Then
        strBuff = left$(strBuff, intPos - 1)
    End If
    
    Global_GetPcName = strBuff
    
End Function
