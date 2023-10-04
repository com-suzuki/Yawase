VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptRYpmf170 
   ClientHeight    =   11010
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptRYpmf170.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptRYpmf170.dsx":000C
End
Attribute VB_Name = "rptRYpmf170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnMeisaiPrint As Boolean

'目　的　　：
'条　件　　：レポートエラー時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：
'更新履歴　：
'
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports.IReturnBool)

    'エラーを表示する
    CancelDisplay = False

End Sub

'目　的　　：
'条　件　　：データがない場合
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：
'更新履歴　：
'
Private Sub ActiveReport_NoData()

    On Error Resume Next

    Call MsgBox("データがありません。", vbOKOnly + vbInformation, "情報")
    Me.Cancel

End Sub

'目　的　　：
'条　件　　：レポート処理を開始する直前
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：
'更新履歴　：
'
Private Sub ActiveReport_ReportStart()
    
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    
    On Error GoTo ActiveReport_ReportStart_Err
       
    '印刷日付
    Me.PrintDay.Text = Format(Now(), "yyyy/mm/dd")
    With adoRecordset1
        '設定マスタ
        strSQL = "{call sp_MT010;1}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            txtCname.Text = IIf(IsNull(.Fields("Company")), "", Trim(.Fields("Company")))
        Else
            txtCname.Text = ""
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Detail.Visible = frmRYpmf170.optPrint(0).Value
    txtOdate.Visible = frmRYpmf170.optPrint(1).Value
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("レポート処理を開始する直前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub PageHeader_Format()
    
    txtBnameStr.Text = Trim(txtBname.Text) & "　様"
    
End Sub
