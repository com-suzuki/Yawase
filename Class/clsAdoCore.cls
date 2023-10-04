VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdoCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_adoConnect As New ADODB.Connection

Private Type typProperty
    Provider As String
    Server As String
    DBName As String
    UID As String
    PWD As String
    LogMess As String
    CommandTimeOut As Integer
End Type
Private m_myProperty As typProperty

Public Enum ProviderName
    adoSQLServer = 1
    adoAccess = 2
End Enum

Private Const ProviderName_Access = "Microsoft.Jet.OLEDB.4.0"
Private Const ProviderName_SQLServer = "SQLOLEDB"

Public Property Get Connection() As ADODB.Connection

    Set Connection = m_adoConnect

End Property

Public Property Let Provider(ByVal Param As ProviderName)

    m_myProperty.Provider = Param

End Property

Public Property Let UID(ByVal Param As String)

    m_myProperty.UID = Param

End Property

Public Property Let PWD(ByVal Param As String)

    m_myProperty.PWD = Param

End Property

Public Property Let DBName(ByVal Param As String)

    m_myProperty.DBName = Param

End Property

Public Property Let Server(ByVal Param As String)

    m_myProperty.Server = Param

End Property

Public Property Let LogMess(ByVal Param As String)

    m_myProperty.LogMess = Param

End Property

Public Property Let CommandTimeOut(ByVal Param As Integer)

    m_myProperty.CommandTimeOut = Param

End Property

Public Function Connect() As Boolean

    On Error GoTo Err:
    
    Connect = False
    
    'データベース接続
    If Trim(m_myProperty.Provider) = "" Then Exit Function
    
    Select Case m_myProperty.Provider
        Case ProviderName.adoSQLServer:    'SQL-Server接続
            With m_adoConnect
                .ConnectionString = "UID=" & m_myProperty.UID & ";" & _
                "PWD=" & m_myProperty.PWD & ";" & _
                "Database=" & m_myProperty.DBName & ";" & _
                "Server=" & m_myProperty.Server & ";"
                .CommandTimeOut = m_myProperty.CommandTimeOut
                .ConnectionTimeout = 60
'                .IsolationLevel = adXactCursorStability
'                .CursorLocation = adUseServer
                .IsolationLevel = adXactChaos
                .CursorLocation = adUseClient
                .Mode = adModeReadWrite
                .Provider = ProviderName_SQLServer
                .Open
            End With
        Case ProviderName.adoAccess:       'Access接続
            With m_adoConnect
                .Open "Provider=" & ProviderName_Access & ";" & _
                      "Data Source=" & m_myProperty.DBName & ";"
            End With
    End Select
    
    Connect = True
    
    Exit Function

Err:

    Dim adoErr As ADODB.Error
    Dim strErr As String
    For Each adoErr In m_adoConnect.Errors
        Call MsgBox("Connect_Error:" & vbCrLf & CStr(adoErr.Number) & adoErr.Description & vbCr & _
                "   (Source: " & adoErr.Source & ")" & vbCr & _
                "   (SQL State: " & adoErr.SQLState & ")" & vbCr & _
                "   (NativeError: " & adoErr.NativeError & ")" _
                , vbOKOnly + vbCritical, "clsSQLCore Error")
    Next adoErr
      
End Function

Public Sub adoClose()
    
    If m_adoConnect.State <> adStateClosed Then m_adoConnect.Close
    Set m_adoConnect = Nothing

End Sub

Private Sub Class_Initialize()

    With m_myProperty
        .UID = ""
        .PWD = ""
        .DBName = ""
        .Server = ""
        .LogMess = ""
        .CommandTimeOut = 0
    End With
    
End Sub

Private Sub Class_Terminate()
    
    If m_adoConnect.State <> adStateClosed Then m_adoConnect.Close
    Set m_adoConnect = Nothing
    
End Sub

