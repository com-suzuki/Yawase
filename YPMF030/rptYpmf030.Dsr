VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf030 
   ClientHeight    =   11010
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf030.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptYpmf030.dsx":000C
End
Attribute VB_Name = "rptYpmf030"
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
    
    On Error GoTo ActiveReport_ReportStart_Err
       
    '印刷日付
    Me.PrintDay.Text = Format(Now(), "yyyy/mm/dd")
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("レポート処理を開始する直前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub GroupHeader1_Format()

    On Error GoTo GroupHeader1_Format_Err

    If txtDiv.Text = CStr(TIKU_DIV_OFF) Then
        txtTiku.Text = "市外"
    ElseIf txtDiv.Text = CStr(TIKU_DIV_ON) Then
        txtTiku.Text = "市内"
    Else
        txtTiku.Text = ""
    End If

    '送金
    If Trim(txtSoukin.Text) = "1" Then
        txtSoukinMsg.Text = "※送金する"
    Else
        txtSoukinMsg.Text = ""
    End If

    Exit Sub

GroupHeader1_Format_Err:
    
   Call MsgBox("GroupHeader1エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupHeader1_Format_Err")

End Sub
