VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'******************************************************************
'
'   クラス名：ログファイルクラス
'
'   処理内容：
'
'   前提条件：OSはWindowsNT4.0, Windows2000のみ対応
'
'   作成者　　：株式会社 コム・エンジニアリング　渥美
'   作成年月日：２００３／０７／０９
'   更新履歴　：
'
'******************************************************************

Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Private Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Integer, ByVal wCategory As Integer, ByVal dwEventID As Long, ByVal lpUserSid As Any, ByVal wNumStrings As Integer, ByVal dwDataSize As Long, plpStrings As Long, lpRawData As Any) As Boolean
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Const EVENTLOG_SUCCESS = 0
Private Const EVENTLOG_ERROR_TYPE = 1
Private Const EVENTLOG_WARNING_TYPE = 2
Private Const EVENTLOG_INFORMATION_TYPE = 4
Private Const EVENTLOG_AUDIT_SUCCESS = 8
Private Const EVENTLOG_AUDIT_FAILURE = 10

Private Declare Function GetVersion Lib "kernel32" () As Long

Dim m_blnNT As Boolean          'WindowsNT系OS

'目　的　　：ログの書き込み
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００３／０７／０９
'更新履歴　：
'
Public Sub SetMessage(ByVal strFileName As String, _
                      ByVal strProgram_Name As String, _
                      ByVal strFunction_Name As String, _
                      ByVal strMessage As String, _
                      ByVal intEvent As Integer)
    
    Dim intFileNumber As Integer
    
    On Error GoTo SetMessage_Err
    
    If intEvent = 1 And m_blnNT = True Then
        'ＮＴのイベントログの書き込み
        Call LogNTEvent(vbCrLf & _
                        "関数名：" & strFunction_Name & vbCrLf & _
                        "エラー内容：" & strMessage _
                        , EVENTLOG_ERROR_TYPE, "1001")
    Else
        '未使用のファイル番号を取得
        intFileNumber = FreeFile
        
        'オープン
        Open strFileName For Append Access Read Write Lock Write As intFileNumber
       
        'ファイル書込み
        Write #intFileNumber, CStr(Now()), strProgram_Name, strFunction_Name, strMessage
    
        'クローズ
        Close intFileNumber
    End If

    Exit Sub

SetMessage_Err:

    Close

End Sub

'目　的　　：ＮＴのイベントログの書き込み
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００３／０７／０９
'更新履歴　：
'
Private Sub LogNTEvent(sString As String, iLogType As Integer, iEventID As Long)
    
    Dim bRC As Boolean
    Dim iNumStrings As Integer
    Dim hEventLog As Long
    Dim hMsgs As Long
    Dim cbStringSize As Long
    
    On Error GoTo LogNTEvent_Err
    
    hEventLog = RegisterEventSource("", App.Title)
    cbStringSize = Len(sString) * 2 + 1
    hMsgs = GlobalAlloc(&H40, cbStringSize)
    CopyMemory ByVal hMsgs, ByVal sString, cbStringSize
    
    iNumStrings = 1
    If ReportEvent(hEventLog, _
       iLogType, 0, _
       iEventID, 0&, _
       iNumStrings, cbStringSize, _
       hMsgs, hMsgs) = 0 Then
       MsgBox GetLastError()
    End If
    
    Call GlobalFree(hMsgs)
    DeregisterEventSource (hEventLog)

    Exit Sub

LogNTEvent_Err:

    MsgBox Error$

End Sub

Private Sub Class_Initialize()

    'Windowsのバージョンチェック
    m_blnNT = IIf(GetVersion() And &H80000000, False, True)
    
End Sub

