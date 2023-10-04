VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptMT070_2 
   ClientHeight    =   11010
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptMT070_2.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptMT070_2.dsx":000C
End
Attribute VB_Name = "rptMT070_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    frmPrintDialog.objArPrint.NoData = True
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
    Dim adoRecordset1 As New ADODB.RecordSet
    
    On Error GoTo ActiveReport_ReportStart_Err
        
    With adoRecordset1
        '設定マスタ
        strSQL = "{call sp_MT010;1}"
        .Open strSQL, frmMt070.m_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            txtCaddress.Text = IIf(IsNull(.Fields("CAddres")), "", Trim(.Fields("CAddres")))
            txtCname.Text = IIf(IsNull(.Fields("Company")), "", Trim(.Fields("Company")))
            txtCceo.Text = IIf(IsNull(.Fields("CCeo")), "", Trim(.Fields("CCeo")))
            txtCTel.Text = "TEL " & IIf(IsNull(.Fields("CTel")), "", Trim(.Fields("CTel")))
            txtCfax.Text = "FAX " & IIf(IsNull(.Fields("CFax")), "", Trim(.Fields("CFax")))
            txtCurl.Text = "URL " & IIf(IsNull(.Fields("CUrl")), "", Trim(.Fields("CUrl")))
            txtCyubin.Text = IIf(IsNull(.Fields("CPost")), "", Trim(.Fields("CPost")))
        Else
            txtCaddress.Text = ""
            txtCname.Text = ""
            txtCceo.Text = ""
            txtCTel.Text = ""
            txtCfax.Text = ""
            txtCurl.Text = ""
            txtCyubin.Text = ""
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    '郵便番号
    txtCyubin1.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 1, 1)
    txtCyubin2.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 2, 1)
    txtCyubin3.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 3, 1)
    txtCyubin4.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 4, 1)
    txtCyubin5.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 5, 1)
    txtCyubin6.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 6, 1)
    txtCyubin7.Text = Mid(Trim(Replace(txtCyubin.Text, "-", "")), 7, 1)
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("レポート処理を開始する直前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub Detail_Format()

    On Error GoTo Detail_Format_Err
    
    '郵便番号
    txtYubin1.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 1, 1)
    txtYubin2.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 2, 1)
    txtYubin3.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 3, 1)
    txtYubin4.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 4, 1)
    txtYubin5.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 5, 1)
    txtYubin6.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 6, 1)
    txtYubin7.Text = Mid(Trim(Replace(txtYubin.Text, "-", "")), 7, 1)
    
    '宛名
    If Trim(txtCeo.Text) <> "" And Trim(txtCeo.Text) <> Trim(txtBname.Text) Then
        txtBnameStr.Text = Trim(txtBname.Text) & vbCrLf & Trim(txtCeo.Text) & "　様"
    Else
        txtBnameStr.Text = Trim(txtBname.Text) & "　様"
    End If
    
    '買主コード
    txtBcodeStr.Text = "(" & txtBcode.Text & ")"
    
    Exit Sub
    
Detail_Format_Err:
    
   Call MsgBox("詳細フォーマットエラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Format_Err")

End Sub
