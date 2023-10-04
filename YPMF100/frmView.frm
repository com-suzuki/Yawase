VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmView 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "購入明細"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12930
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   12930
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdExecute 
      Cancel          =   -1  'True
      Caption         =   "一覧印刷"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8865
      TabIndex        =   1
      Top             =   9120
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmView.frx":0CFA
      Height          =   8175
      Left            =   240
      TabIndex        =   3
      Top             =   670
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   14420
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Pnum"
         Caption         =   "受付番号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PnumLine"
         Caption         =   "行"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Iname"
         Caption         =   "植木名称"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Qty"
         Caption         =   "数量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Price"
         Caption         =   "売立金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Ocode"
         Caption         =   "競売番号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Bdiv"
         Caption         =   "発行済"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   "0%"
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Bnum"
         Caption         =   "回数"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2805.166
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   824.882
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmView.frx":0D0F
      Height          =   8175
      Left            =   240
      TabIndex        =   5
      Top             =   660
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   14420
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Onum"
         Caption         =   "注文番号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Line"
         Caption         =   "行"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Iname"
         Caption         =   "植木名称"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Qty"
         Caption         =   "数量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Price"
         Caption         =   "売立金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Bdiv"
         Caption         =   "発行済"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Bnum"
         Caption         =   "回数"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4410.142
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "戻　る"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10905
      TabIndex        =   0
      Top             =   9120
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   8385
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=YAWASESRC"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "YAWASESRC"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "SELECT * FROM DT021 ORDER BY Ocode,Line"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   435
      Left            =   10605
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=YAWASESRC"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "YAWASESRC"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "SELECT * FROM DT031 ORDER BY Bcode,Bnum,Line"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8955
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15796
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "競　売"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "注　文"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblRecordCount1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   9180
      Width           =   4275
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_intBode As Integer
Public m_intBnum As Integer
Public m_strBname As String

'目　的　　：
'条　件　　：印刷クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／０９／０８
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    '印刷用ワーク作成
    If MakePrintWork() = False Then Exit Sub
    '印刷プレビュー
    If ActiveReportPrint(0) = False Then Exit Sub
    
    Exit Sub
    
cmdExecute_Click_Err:

    Call MsgBox("印刷クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Dim strSQL As String
    Dim lngRecordCount As Long

    On Error GoTo Form_Load_Err

    '競売明細データ
    strSQL = "SELECT * FROM DT021" & _
             " WHERE Bcode = " & m_intBode & _
             " AND Bnum = " & m_intBnum & _
             " AND LEFT(Ocode,8) = '" & Global_Get_NumericDay(g_strOdate) & "'" & _
             " ORDER BY Ocode,Line"
    Adodc1.RecordSource = strSQL
    Adodc1.Refresh

    '注文明細データ
    strSQL = "SELECT * FROM DT031" & _
             " WHERE Bcode = " & m_intBode & _
             " AND Bnum = " & m_intBnum & _
             " AND Odate = '" & g_strOdate & "'" & _
             " AND Price <> 0" & _
             " ORDER BY Onum,Line"
    Adodc2.RecordSource = strSQL
    Adodc2.Refresh

    lngRecordCount = Adodc1.Recordset.RecordCount + Adodc2.Recordset.RecordCount
    lblRecordCount1.Caption = lngRecordCount & "件"
    If Adodc2.Recordset.RecordCount > 0 Then
        lblRecordCount1.Caption = lblRecordCount1.Caption & "(注文分 " & Adodc2.Recordset.RecordCount & "件)"
    End If

    Exit Sub

Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")

End Sub

Private Sub TabStrip1_Click()

    If TabStrip1.Tabs(1).Selected = True Then
        DataGrid1.ZOrder 0
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        DataGrid2.ZOrder 0
    End If
    
End Sub

'目　的　　：ActiveReportの印刷
'条　件　　：
'結　果　　：
'引　数　　：0:プレビュー 1:印刷
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／０８
'更新履歴　：
'
Private Function ActiveReportPrint(intFlg As Integer) As Boolean
    
    Dim objRpt As New rptYpmf100
    Dim objArPrint As New clsArPrint
    
    On Error GoTo ActiveReportPrint_Err
    
    ActiveReportPrint = False
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "買主購入明細"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .Caption = "買主購入明細"
        If .PrintActiveReport(intFlg) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End With

    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    ActiveReportPrint = True
    
    Exit Function
    
ActiveReportPrint_Err:

    ActiveReportPrint = False
    Screen.MousePointer = vbDefault
    Call MsgBox("実行クリックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReportPrint_Err")
    
End Function

'目　的　　：印刷用ワーク作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／０９／０８
'更新履歴　：
'
Private Function MakePrintWork() As Boolean

    Dim strSQL As String
    Dim wkRecordset As New ADODB.Recordset
    
    On Error GoTo MakePrintWork_Err
    
    MakePrintWork = False
    
    Screen.MousePointer = vbHourglass
    
    If Adodc1.Recordset.BOF And Adodc1.Recordset.EOF And Adodc2.Recordset.BOF And Adodc2.Recordset.EOF Then
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
        GoTo MakePrintWork_Exit:
    End If
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF100"
    g_clsAdoAccess.Connection.Execute strSQL

    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF100"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
        Adodc1.Recordset.MoveFirst
        Do While Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False
            wkRecordset.AddNew
            wkRecordset.Fields("Odate") = g_strOdate
            wkRecordset.Fields("Bcode") = Adodc1.Recordset("Bcode")
            wkRecordset.Fields("Bname") = Trim(m_strBname)
            wkRecordset.Fields("Kbn") = "競売"
            wkRecordset.Fields("Num") = Adodc1.Recordset("Pnum")
            wkRecordset.Fields("Line") = Adodc1.Recordset("PnumLine")
            wkRecordset.Fields("Iname") = Adodc1.Recordset("Iname")
            wkRecordset.Fields("Qty") = Adodc1.Recordset("Qty")
            If Adodc1.Recordset("Price") = 0 Or Adodc1.Recordset("Qty") = 0 Then
                wkRecordset.Fields("Price1") = 0
            Else
                wkRecordset.Fields("Price1") = Fix(CCur(Adodc1.Recordset("Price")) / CCur(Adodc1.Recordset("Qty")))
            End If
            wkRecordset.Fields("Price") = Adodc1.Recordset("Price")
            wkRecordset.Fields("Ocode") = Adodc1.Recordset("Ocode")
            wkRecordset.Fields("UCount") = 1
            wkRecordset.Update
            
            Adodc1.Recordset.MoveNext
        Loop
        Adodc1.Recordset.MoveFirst
    End If
       
    If Adodc2.Recordset.BOF = False And Adodc2.Recordset.EOF = False Then
        Adodc2.Recordset.MoveFirst
        Do While Adodc2.Recordset.BOF = False And Adodc2.Recordset.EOF = False
            wkRecordset.AddNew
            wkRecordset.Fields("Odate") = g_strOdate
            wkRecordset.Fields("Bcode") = Adodc2.Recordset("Bcode")
            wkRecordset.Fields("Bname") = Trim(m_strBname)
            wkRecordset.Fields("Kbn") = "注文"
            wkRecordset.Fields("Num") = Adodc2.Recordset("Onum")
            wkRecordset.Fields("Line") = Adodc2.Recordset("Line")
            wkRecordset.Fields("Iname") = Adodc2.Recordset("Iname")
            wkRecordset.Fields("Qty") = Adodc2.Recordset("Qty")
            wkRecordset.Fields("Price1") = Adodc2.Recordset("Price1")
            wkRecordset.Fields("Price") = Adodc2.Recordset("Price")
            wkRecordset.Fields("Ocode") = ""
            wkRecordset.Fields("CCount") = 1
            wkRecordset.Update
            
            Adodc2.Recordset.MoveNext
        Loop
        Adodc2.Recordset.MoveFirst
    End If
       
    wkRecordset.Requery
    wkRecordset.Close
    
    MakePrintWork = True
    
MakePrintWork_Exit:
    
    Screen.MousePointer = vbDefault
    
    Exit Function

MakePrintWork_Cancel:

    GoTo MakePrintWork_Exit:

MakePrintWork_Err:

    MakePrintWork = False
    Call MsgBox("印刷ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakePrintWork_Err")
    GoTo MakePrintWork_Exit:

End Function
