VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdoRecordCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_adoConnect As ADODB.Connection
Private m_adoRecordset As New ADODB.Recordset
Private m_varKeyName As Variant
Private m_varKeyValue As Variant
Private m_strTableName As String

Private m_strMoveNextSQL As String
Private m_strMovePrevious As String

Public Property Get Connection() As ADODB.Connection

    Set Connection = m_adoConnect

End Property

Public Property Let Connection(ByVal Param As ADODB.Connection)

    Set m_adoConnect = Param

End Property

Public Property Get Recordset() As ADODB.Recordset

    Set Recordset = m_adoRecordset

End Property

Public Property Let Recordset(ByVal Param As ADODB.Recordset)

    Set m_adoRecordset = Param

End Property

Public Property Get KeyName() As Variant

    KeyName = m_varKeyName

End Property

Public Property Let KeyName(ByVal Param As Variant)

    m_varKeyName = Param

End Property

Public Property Get KeyValue() As Variant

    KeyValue = m_varKeyValue

End Property

Public Property Let KeyValue(ByVal Param As Variant)

    m_varKeyValue = Param

End Property

Public Property Get TableName() As String

    TableName = m_strTableName

End Property

Public Property Let TableName(ByVal Param As String)

    m_strTableName = Param

End Property

Public Sub MoveFirst()

    On Error GoTo MoveFirst_Err

    With m_adoRecordset
        .Open m_strTableName, m_adoConnect, adOpenKeyset, adLockReadOnly
        If Not .BOF And Not .EOF Then
            .MoveFirst
        End If
    End With
    
    Exit Sub
    
MoveFirst_Err:
    
    Call MsgBox("レコード移動エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MoveFirst_Err")
    
End Sub

Public Sub MoveLast()

    On Error GoTo MoveLast_Err
    
    With m_adoRecordset
        .Open m_strTableName, m_adoConnect, adOpenKeyset, adLockReadOnly
        If Not .BOF And Not .EOF Then
            .MoveLast
        End If
    End With

    Exit Sub

MoveLast_Err:
    
    Call MsgBox("レコード移動エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MoveLast_Err")

End Sub

Public Sub MoveNext()

    Dim strSQL As String
    Dim strWhere As String
    Dim strOrder As String
    Dim intIndex1 As Integer

    On Error GoTo MoveNext_Err

    If IsNull(m_varKeyName) = True Then Exit Sub
    If IsNull(m_varKeyValue) = True Then Exit Sub
    If UBound(m_varKeyName) <> UBound(m_varKeyValue) Then Exit Sub

    'SQL文作成
    strSQL = "SELECT * FROM " & m_strTableName
    
    strWhere = ""
    strOrder = ""
    For intIndex1 = 0 To UBound(m_varKeyName)
        'WHERE句
        If strWhere <> "" Then strWhere = strWhere & " AND "
        Select Case TypeName(m_varKeyValue(intIndex1))
            '数値型
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Boolean":
                strWhere = strWhere & m_varKeyName(intIndex1) & " >= " & m_varKeyValue(intIndex1)
            '文字型
            Case "String":
                strWhere = strWhere & m_varKeyName(intIndex1) & " >= '" & m_varKeyValue(intIndex1) & "'"
            Case Else
                strWhere = strWhere & m_varKeyName(intIndex1) & " >= '" & m_varKeyValue(intIndex1) & "'"
        End Select
        
        'ORDER句
        If strOrder <> "" Then strOrder = strOrder & ","
        strOrder = strOrder & m_varKeyName(intIndex1) & " ASC"
    Next intIndex1
    
    strSQL = strSQL & " WHERE " & strWhere & " ORDER BY " & strOrder
                
    With m_adoRecordset
        .Open strSQL, m_adoConnect, adOpenKeyset, adLockReadOnly
        If Not .EOF And Not .BOF Then
            .MoveNext
        Else
            .MoveLast
        End If
    End With
                
    Exit Sub
                
MoveNext_Err:
                
    Call MsgBox("レコード移動エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MoveNext_Err")
                
End Sub

Public Sub MovePrevious()

    Dim strSQL As String
    Dim strWhere As String
    Dim strOrder As String
    Dim intIndex1 As Integer

    On Error GoTo MovePrevious_Err

    If IsNull(m_varKeyName) = True Then Exit Sub
    If IsNull(m_varKeyValue) = True Then Exit Sub
    If UBound(m_varKeyName) <> UBound(m_varKeyValue) Then Exit Sub

    'SQL文作成
    strSQL = "SELECT * FROM " & m_strTableName
    
    strWhere = ""
    strOrder = ""
    For intIndex1 = 0 To UBound(m_varKeyName)
        'WHERE句
        If strWhere <> "" Then strWhere = strWhere & " AND "
        Select Case TypeName(m_varKeyValue(intIndex1))
            '数値型
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Boolean":
                strWhere = strWhere & m_varKeyName(intIndex1) & " <= " & m_varKeyValue(intIndex1)
            '文字型
            Case "String":
                strWhere = strWhere & m_varKeyName(intIndex1) & " <= '" & m_varKeyValue(intIndex1) & "'"
            Case Else
                strWhere = strWhere & m_varKeyName(intIndex1) & " <= '" & m_varKeyValue(intIndex1) & "'"
        End Select
        
        'ORDER句
        If strOrder <> "" Then strOrder = strOrder & ","
        strOrder = strOrder & m_varKeyName(intIndex1) & " DESC"
    Next intIndex1
    
    strSQL = strSQL & " WHERE " & strWhere & " ORDER BY " & strOrder
                
    With m_adoRecordset
        .Open strSQL, m_adoConnect, adOpenKeyset, adLockReadOnly
        If Not .EOF And Not .BOF Then
            .MoveNext
        Else
            .MoveLast
        End If
    End With
                
    Exit Sub
                
MovePrevious_Err:
                
    Call MsgBox("レコード移動エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MovePrevious_Err")

End Sub

Private Sub Class_Initialize()

    Set m_adoRecordset = Nothing
    m_varKeyName = Null
    m_varKeyValue = Null
    m_strTableName = ""

End Sub
