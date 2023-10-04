VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf050 
   ClientHeight    =   11010
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf050.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptYpmf050.dsx":000C
End
Attribute VB_Name = "rptYpmf050"
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
    If Trim(txtIdiv.Text) <> "" And Trim(txtIdiv.Text) <> "0" And Trim(txtPrice.Text) = "0" Then
        txtPrice.Visible = False
        txtMsg.Visible = True
        txtMsg.Text = "(合 算)"
    End If
    
    If txtLine.Text = "0" Then
        txtLine.Text = ""
    End If
    
    Exit Sub

Detail_Format_Err:

   Call MsgBox("明細フォーマットエラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Format_Err")

End Sub

Private Sub GroupFooter1_Format()

    On Error GoTo GroupFooter1_Format_Err

'    If Trim(txtCflg.Text) = "1" Then
'        lblCflg.Visible = True
'    Else
'        lblCflg.Visible = False
'    End If

    If txtNum.Text <> "" And txtNum.Text <> "0" And txtNum.Text <> "1" Then
        txtNumMsg.Text = "精算回数：" & txtNum.Text
        txtNumMsg.Visible = True
    Else
        txtNumMsg.Visible = False
    End If

    lblTotal.Caption = Format(CDbl(txtSubtotal.DataValue) + CDbl(txtKeep.DataValue) + CDbl(txtCharge.DataValue), "###,###,##0")
    

    Exit Sub

GroupFooter1_Format_Err:

   Call MsgBox("グループフッターフォーマット時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupFooter1_Format_Err")

End Sub

Private Sub GroupFooter3_Format()
    
    If txtMishu.Text = "0" Then
        GroupFooter3.Visible = False
    Else
        GroupFooter3.Visible = True
    End If

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

    If frmYpmf050.chkRePrint.Value = 1 Then
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
    
    txtBcode.Text = txtBcodeFooter.Text
    txtBname.Text = txtBnameFooter.Text

    Exit Sub

PageHeader_BeforePrint_Err:

   Call MsgBox("ページヘッダー印刷前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageHeader_BeforePrint_Err")

End Sub

