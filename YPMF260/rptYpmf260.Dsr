VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf260 
   ClientHeight    =   11115
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf260.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptYpmf260.dsx":000C
End
Attribute VB_Name = "rptYpmf260"
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
       
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("レポート処理を開始する直前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub Detail_Format()

    On Error GoTo Detail_Format_Err
       
    If frmYpmf260.optDiv(1).Value = True Then
        txtUriage_Kingaku.Visible = False
        LineUriage_Kingaku2.Visible = False
    Else
        txtUriage_Kingaku.Visible = True
        LineUriage_Kingaku2.Visible = True
    End If
       
    Exit Sub
    
Detail_Format_Err:
    
   Call MsgBox("明細エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "Detail_Format_Err")

End Sub

Private Sub GroupFooter1_Format()

    On Error GoTo GroupFooter1_Format_Err
       
    '小計タイトルがない場合は表示しない
    If Trim(txtSubtotal_Name.Text) = "" Then
        GroupFooter1.Visible = False
    Else
        GroupFooter1.Visible = True
    End If
       
    If frmYpmf260.optDiv(1).Value = True Then
        txtUriage_Kingaku_Subtotal.Visible = False
        LineUriage_Kingaku3.Visible = False
    Else
        txtUriage_Kingaku_Subtotal.Visible = True
        LineUriage_Kingaku3.Visible = True
    End If
       
    Exit Sub
    
GroupFooter1_Format_Err:
    
   Call MsgBox("グループフッターエラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "GroupFooter1_Format_Err")
  
End Sub

Private Sub PageHeader_Format()

    On Error GoTo PageHeader_Format_Err
       
    If frmYpmf260.optDiv(1).Value = True Then
        lblUriage_Kingaku.Visible = False
        LineUriage_Kingaku1.Visible = False
    Else
        lblUriage_Kingaku.Visible = True
        LineUriage_Kingaku1.Visible = True
    End If
       
    Exit Sub
    
PageHeader_Format_Err:
    
   Call MsgBox("ページヘッダーエラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "PageHeader_Format_Err")
  
End Sub

Private Sub ReportFooter_Format()

    On Error GoTo ReportFooter_Format_Err
       
    txtTokki.Text = frmYpmf260.txtTokki.Text
       
    If frmYpmf260.optDiv(1).Value = True Then
        txtUriage_Kingaku_Total.Visible = False
        LineUriage_Kingaku4.Visible = False
    Else
        txtUriage_Kingaku_Total.Visible = True
        LineUriage_Kingaku4.Visible = True
    End If
       
    Exit Sub
    
ReportFooter_Format_Err:
    
   Call MsgBox("レポートフッターエラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ReportFooter_Format_Err")
  
End Sub
