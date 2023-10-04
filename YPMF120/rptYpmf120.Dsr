VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf120 
   ClientHeight    =   11115
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf120.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptYpmf120.dsx":000C
End
Attribute VB_Name = "rptYpmf120"
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
       
    '印刷日付
    Me.PrintDay.Text = Format(Now(), "yyyy/mm/dd")
    '改ページ制御
    GroupFooter1.NewPage = ddNPAfter
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("レポート処理を開始する直前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub GroupFooter1_Format()
    
    On Error GoTo GroupFooter1_Format_Err
       
    '改ページ制御
    If Me.DataControl1.Recordset.EOF = True Then
        GroupFooter1.NewPage = ddNPNone
    End If
    
    If Trim(txtGTotal1.Text) <> "" And Trim(txtNyukin1.Text) <> "" Then
       txtZandaka1.Text = CCur(txtGTotal1.Text) - CCur(txtNyukin1.Text)
       txtZandaka1.Text = Format(txtZandaka1.Text, "#,##0")
    End If
    
    Exit Sub
    
GroupFooter1_Format_Err:
    
   Call MsgBox("グループフッターフォーマット時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupFooter1_Format_Err")
  
End Sub

Private Sub GroupFooter2_Format()
    
    On Error GoTo GroupFooter2_Format_Err
       
    If Trim(txtGtotal.Text) <> "" And Trim(txtNyukin.Text) <> "" Then
       txtZandaka.Text = CCur(txtGtotal.Text) - CCur(txtNyukin.Text)
       txtZandaka.Text = Format(txtZandaka.Text, "#,##0")
    End If
       
    Exit Sub
    
GroupFooter2_Format_Err:
    
   Call MsgBox("グループフッターフォーマット時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupFooter2_Format_Err")
       
End Sub

Private Sub PageHeader_Format()
     
    On Error GoTo PageHeader_Format_Err
       
    If txtDiv.Text = "0" Then
        lblTitle.Visible = True
        txtOdateTitle.Visible = True
        lineTitle.Visible = True
    Else
        lblTitle.Visible = False
        txtOdateTitle.Visible = False
        lineTitle.Visible = False
    End If
           
    Exit Sub
    
PageHeader_Format_Err:
    
   Call MsgBox("ページヘッダーフォーマット時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageHeader_Format_Err")
  

End Sub

Private Sub ReportFooter_Format()
        
    On Error GoTo ReportFooter_Format_Err
       
    If Trim(txtGtotal_Total.Text) <> "" And Trim(txtNyukin_total.Text) <> "" Then
       txtZandaka_Total.Text = CCur(txtGtotal_Total.Text) - CCur(txtNyukin_total.Text)
       txtZandaka_Total.Text = Format(txtZandaka_Total.Text, "#,##0")
    End If
       
    Exit Sub
    
ReportFooter_Format_Err:
    
   Call MsgBox("レポートフッターフォーマット時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ReportFooter_Format_Err")
    
End Sub
