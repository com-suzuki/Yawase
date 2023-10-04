VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typProperty
    RegKey As String            'レジストリのキー
    Bin As String
    LDatabase As String
    LDBName As String
    Log As String
    Server As String
    DBName As String
    UID As String
    PWD As String
    CommandTimeOut As Integer
    DSN As String
    BackUpDrive As String
    DownLoadPath As String
End Type
Private m_myProperty As typProperty

'レジスト操作用 WindowsAPI
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_RESOURCE_LIST = 8
Private Const REG_SZ = 1

Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0&
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private hKey As Long
Private lpSubKey As String
Private phKey As Long

'目　的　　：キーのクローズ
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング 渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function ComregClose() As Boolean

    Dim lErr As Long
    
    On Error GoTo ComregClose_Err
    
    lErr = RegCloseKey(phKey)
    If lErr = ERROR_SUCCESS Then
        ComregClose = True
    Else
        ComregClose = False
    End If
    
    Exit Function

ComregClose_Err:
    
    ComregClose = False
    Call MsgBox("レジストリークローズエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ComregClose_Err")

End Function

'目　的　　：キー値の作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング 渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function ComregCreate(ByRef lpSubKey As String) As Boolean
    
    Dim lErr As Long
    Dim lpClass As String
    Dim lpdwDisposition As Long
    Dim seq As SECURITY_ATTRIBUTES
    
    On Error GoTo ComregCreate_Err
    
    hKey = HKEY_LOCAL_MACHINE
    lErr = RegCreateKeyEx(hKey, lpSubKey, 0, lpClass, 0, KEY_ALL_ACCESS, seq, phKey, lpdwDisposition)
    If lErr = ERROR_SUCCESS Then
        ComregCreate = True
    Else
        ComregCreate = False
    End If
    
    Exit Function
    
ComregCreate_Err:

    ComregCreate = False
    Call MsgBox("レジストリーサブキー作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ComregCreate_Err")

End Function

'目　的　　：キーのオープン
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング 渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function ComregOpen(ByRef lpStr As String)

    Dim lErr As Long
    Dim ulOption As Long
    Dim samDesired As Long
    
    On Error GoTo ComregOpen_Err
    
    hKey = HKEY_LOCAL_MACHINE
    lpSubKey = lpStr
    ulOption = 0
    
    lErr = RegOpenKeyEx(hKey, lpStr, 0, KEY_ALL_ACCESS, phKey)
    If lErr = ERROR_SUCCESS Then
        ComregOpen = True
    Else
        ComregOpen = False
    End If
    
    Exit Function
    
ComregOpen_Err:
    
    ComregOpen = False
    Call MsgBox("レジストリーオープンエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ComregOpen_Err")

End Function

'目　的　　：キー値の検索
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング 渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function ComregQueryText(ByRef lpValueName As String) As String

    Dim lErr As Long
    Dim lpReserved As Long
    Dim lpType As Long
    Dim lpData As String * 255
    Dim lpcbData As Long
    
    On Error GoTo ComregQueryText_Err
    
    lpReserved = 0
    ComregQueryText = ""
    
    lpcbData = Len(lpData) - 1
    lErr = RegQueryValueEx(phKey, lpValueName, 0, lpType, ByVal lpData, lpcbData)
    If lErr = ERROR_SUCCESS Then
        ComregQueryText = Left(lpData, InStr(lpData, Chr(0)) - 1)
    End If

    Exit Function

ComregQueryText_Err:

    ComregQueryText = ""
    Call MsgBox("レジストリー検索エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ComregQueryText_Err")
 
End Function

'目　的　　：キー値の格納（文字列）
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング 渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function ComregSetText(ByRef lpValueName As String, ByRef lpData As String) As Boolean

    Dim lErr As Long
    Dim Reserved As Long
    Dim dwType As Long
    Dim cbData As Long

    On Error GoTo ComregSetText_Err
    
    dwType = REG_SZ
    Reserved = 0
'    cbData = Len(lpData)
    cbData = LenB(lpData)
    lErr = RegSetValueEx(phKey, lpValueName, Reserved, dwType, ByVal lpData, cbData + 1)
    If lErr = ERROR_SUCCESS Then
        ComregSetText = True
    Else
        ComregSetText = False
    End If
    
    Exit Function
    
ComregSetText_Err:

    ComregSetText = False
    Call MsgBox("レジストリー書き込みエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ComregSetText_Err")
    
End Function

Private Sub Class_Initialize()

    With m_myProperty
        .RegKey = ""
        .Bin = ""
        .LDatabase = ""
        .LDBName = ""
        .Log = ""
        .Server = ""
        .DBName = ""
        .UID = ""
        .PWD = ""
        .CommandTimeOut = 0
        .DSN = ""
        .DownLoadPath = ""
    End With

End Sub

Private Sub Class_Terminate()
        
    '処理なし
   
End Sub

Public Function ReadReg() As Boolean

    Dim strRegKey As String

    On Error GoTo ReadReg_Err

    ReadReg = False
        
    If Trim(m_myProperty.RegKey) = "" Then Exit Function
    
    'レジストリオープン
    If ComregOpen(m_myProperty.RegKey) = False Then Exit Function
        
    'キーの検索
    m_myProperty.Bin = ComregQueryText("Bin")
    m_myProperty.LDatabase = ComregQueryText("LDatabase")
    m_myProperty.LDBName = ComregQueryText("LDBName")
    m_myProperty.Log = ComregQueryText("Log")
    m_myProperty.Server = ComregQueryText("Server")
    m_myProperty.DBName = ComregQueryText("DBName")
    m_myProperty.UID = ComregQueryText("UID")
    m_myProperty.PWD = ComregQueryText("PWD")
    m_myProperty.CommandTimeOut = IIf(IsNumeric(ComregQueryText("CommandTimeOut")) = True, ComregQueryText("CommandTimeOut"), 0)
    m_myProperty.DSN = ComregQueryText("DSN")
    m_myProperty.BackUpDrive = ComregQueryText("BackUpDrive")
    m_myProperty.DownLoadPath = ComregQueryText("DownLoadPath")
    
    'レジストリクローズ
    If ComregClose() = False Then Exit Function

    ReadReg = True

    Exit Function

ReadReg_Err:

    ReadReg = False
    Call MsgBox("レジストリー読み込みエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ReadReg_Err")
    
End Function

Public Function WriteReg() As Boolean

    Dim strRegKey As String

    On Error GoTo WriteReg_Err

    WriteReg = False
        
    If Trim(m_myProperty.RegKey) = "" Then Exit Function
    
    'レジストリオープン
    If ComregOpen(m_myProperty.RegKey) = False Then
        'キー作成
        If ComregCreate(m_myProperty.RegKey) = False Then
            Exit Function
        End If
    End If
        
    'キーの保存
    Call ComregSetText("Bin", CStr(m_myProperty.Bin))
    Call ComregSetText("LDatabase", CStr(m_myProperty.LDatabase))
    Call ComregSetText("LDBName", CStr(m_myProperty.LDBName))
    Call ComregSetText("Log", CStr(m_myProperty.Log))
    Call ComregSetText("Server", CStr(m_myProperty.Server))
    Call ComregSetText("DBName", CStr(m_myProperty.DBName))
    Call ComregSetText("UID", CStr(m_myProperty.UID))
    Call ComregSetText("PWD", CStr(m_myProperty.PWD))
    Call ComregSetText("CommandTimeOut", CStr(m_myProperty.CommandTimeOut))
    Call ComregSetText("DSN", CStr(m_myProperty.DSN))
    Call ComregSetText("BackUpDrive", CStr(m_myProperty.BackUpDrive))
    Call ComregSetText("DownLoadPath", CStr(m_myProperty.DownLoadPath))
    
    'レジストリクローズ
    If ComregClose() = False Then Exit Function

    WriteReg = True

    Exit Function

WriteReg_Err:

    WriteReg = False
    Call MsgBox("レジストリー書き込みエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "WriteReg_Err")
    
End Function

Public Property Get RegKey() As String

    RegKey = m_myProperty.RegKey

End Property

Public Property Let RegKey(ByVal Param As String)

    m_myProperty.RegKey = Param

End Property

Public Property Get Server() As String

    Server = m_myProperty.Server

End Property

Public Property Let Server(ByVal Param As String)

    m_myProperty.Server = Param

End Property

Public Property Get DBName() As String

    DBName = m_myProperty.DBName

End Property

Public Property Let DBName(ByVal Param As String)

    m_myProperty.DBName = Param

End Property

Public Property Get UID() As String

    UID = m_myProperty.UID

End Property

Public Property Let UID(ByVal Param As String)

    m_myProperty.UID = Param

End Property

Public Property Get PWD() As String

    PWD = m_myProperty.PWD

End Property

Public Property Let PWD(ByVal Param As String)

    m_myProperty.PWD = Param

End Property

Public Property Get CommandTimeOut() As Integer

    CommandTimeOut = m_myProperty.CommandTimeOut

End Property

Public Property Let CommandTimeOut(ByVal Param As Integer)

    m_myProperty.CommandTimeOut = Param

End Property

Public Property Get Bin() As String

    Bin = m_myProperty.Bin

End Property

Public Property Let Bin(ByVal Param As String)

    m_myProperty.Bin = Param

End Property

Public Property Get LDatabase() As String

    LDatabase = m_myProperty.LDatabase

End Property

Public Property Let LDatabase(ByVal Param As String)

    m_myProperty.LDatabase = Param

End Property

Public Property Get LDBName() As String

    LDBName = m_myProperty.LDBName

End Property

Public Property Let LDBName(ByVal Param As String)

    m_myProperty.LDBName = Param

End Property

Public Property Get Log() As String

    Log = m_myProperty.Log

End Property

Public Property Let Log(ByVal Param As String)

    m_myProperty.Log = Param

End Property

Public Property Get DSN() As String

    DSN = m_myProperty.DSN

End Property

Public Property Let DSN(ByVal Param As String)

    m_myProperty.DSN = Param

End Property

Public Property Get BackUpDrive() As String

    BackUpDrive = m_myProperty.BackUpDrive

End Property

Public Property Let BackUpDrive(ByVal Param As String)

    m_myProperty.BackUpDrive = Param

End Property

Public Property Get DownLoadPath() As String

    DownLoadPath = m_myProperty.DownLoadPath

End Property

Public Property Let DownLoadPath(ByVal Param As String)

    m_myProperty.DownLoadPath = Param

End Property


