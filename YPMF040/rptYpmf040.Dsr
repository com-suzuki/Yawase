VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf040 
   ClientHeight    =   14115
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   19080
   Icon            =   "rptYpmf040.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   33655
   _ExtentY        =   24897
   SectionData     =   "rptYpmf040.dsx":000C
End
Attribute VB_Name = "rptYpmf040"
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

    On Error GoTo ActiveReport_ReportStart_Err

    Me.GroupHeader2.NewPage = ddNPNone

    Exit Sub

ActiveReport_ReportStart_Err:

   Call MsgBox("レポート処理を開始する直前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")

End Sub

Private Sub Detail_Format()
    
    On Error GoTo Detail_Format_Err

    txtMsg.Visible = False
    txtPrice.Visible = True
    If Trim(txtIdiv.Text) <> "" And Trim(txtIdiv.Text) <> "0" And Trim(txtIdiv.Text) <> Trim(txtLine.Text) Then
        txtPrice.Visible = False
        txtMsg.Visible = True
        txtMsg.Text = "(�ａF" & txtIdiv.Text & "に合算)"
    End If
    If Trim(txtResult.Text) <> "" And Trim(txtResult.Text) <> "0" Then
        txtPrice.Visible = False
        txtMsg.Visible = True
        txtMsg.Text = "(競売不成立)"
    End If
    
    Exit Sub

Detail_Format_Err:

   Call MsgBox("明細フォーマットエラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Format_Err")

End Sub

Private Sub GroupFooter1_Format()

    '202308インボイス対応
    lblTotal.Caption = Format(txtTotal.DataValue - (txtChargeDisplay.DataValue) - CDbl(txtRrate.DataValue) - CDbl(txtEf.DataValue) - CDbl(txtKeep.DataValue), "###,###,###,##0")
    
End Sub

'目　的　　：
'条　件　　：グループフッター印刷後
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：
'更新履歴　：
'
Private Sub GroupHeader2_AfterPrint()

    On Error GoTo GroupHeader2_AfterPrint_Err

    Me.GroupHeader2.NewPage = ddNPBefore

    Exit Sub

GroupHeader2_AfterPrint_Err:

   Call MsgBox("グループフッター印刷後エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupHeader2_AfterPrint_Err")

End Sub

'目　的　　：
'条　件　　：ページフッターフォーマット時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：
'更新履歴　：
'
Private Sub PageFooter_Format()

    On Error GoTo PageFooter_Format_Err

    If frmYpmf040.chkRePrint.Value = 1 Then
        txtItime.Visible = True
        lblItime.Visible = True
    End If

    Exit Sub

PageFooter_Format_Err:

   Call MsgBox("ページフッターフォーマット時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageFooter_Format_Err")

End Sub

Private Sub PageHeader_BeforePrint()
    
    On Error GoTo PageHeader_BeforePrint_Err
    
    txtPnum.Text = txtPnumFooter.Text
    txtSname.Text = txtSnameFooter.Text
    '202308インボイス対応
    txtTnum.Text = txtTnumFooter.Text

    Exit Sub

PageHeader_BeforePrint_Err:

   Call MsgBox("ページヘッダー印刷前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageHeader_BeforePrint_Err")

End Sub
