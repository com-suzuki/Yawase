VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} rptYpmf130 
   ClientHeight    =   11115
   ClientLeft      =   -3495
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "rptYpmf130.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptYpmf130.dsx":000C
End
Attribute VB_Name = "rptYpmf130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngCount As Long

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

Private Sub ActiveReport_FetchData(eof As Boolean)

    Dim CurrentValue(1) As Variant
    Static BeforeValue(1) As Variant

    On Error GoTo ActiveReport_FetchData_Err
       
    CurrentValue(1) = Me.Fields("Pnum").Value
    
    '前回値と比較
    If CurrentValue(1) = BeforeValue(1) Then
        Me.Fields("Pnum").Value = ""
        Me.Fields("Sname").Value = ""
    End If
    
    BeforeValue(1) = CurrentValue(1)
    
    Exit Sub
    
ActiveReport_FetchData_Err:
    
   Call MsgBox("データ取得時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_FetchData_Err")
    

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
    m_lngCount = 0
    
    Exit Sub
    
ActiveReport_ReportStart_Err:
    
   Call MsgBox("レポート処理を開始する直前エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "ActiveReport_ReportStart_Err")
    
End Sub

Private Sub Detail_Format()

    m_lngCount = m_lngCount + 1
    
End Sub

Private Sub ReportFooter_Format()

    txtCount.Text = CStr(m_lngCount)

End Sub
