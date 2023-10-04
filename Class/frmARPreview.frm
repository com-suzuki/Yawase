VERSION 5.00
Object = "{E95678BE-E45E-471F-9287-59E8911E479E}#1.5#0"; "ArViewer15j.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmARPreview 
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   Icon            =   "frmARPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   9285
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   420
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer arv 
      Height          =   7665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   13520
      SectionData     =   "frmARPreview.frx":0CFA
   End
End
Attribute VB_Name = "frmARPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_objArPrint As clsArPrint

Const TOOLBAR_ADD01 = "ページ設定"
Const TOOLBAR_ADD02 = "PDF出力"
Const TOOLBAR_ADD03 = "閉じる"

Private Sub arv_PrintAborted()

    Unload Me
    
End Sub

Private Sub arv_ToolbarClick(ByVal tool As DDActiveReportsViewerCtl.DDTool)

    Select Case tool.Caption
        Case TOOLBAR_ADD01: 'ページ設定
            Set frmARPage.m_objArPrint = m_objArPrint
            frmARPage.Show vbModal
        Case TOOLBAR_ADD02: 'PDF出力
            Call OutPutPDF
        Case TOOLBAR_ADD03: 'フォームを閉じる
            Unload Me
        Case Else
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    With arv.ToolBar
        'ツールバーの一部のアイテムを非表示にします。
        .Tools(0).Visible = False    '見出し
        .Tools(12).Visible = False   '戻る
        .Tools(13).Visible = False   '進む
    
        'ツールバーの追加
        .Tools.Add TOOLBAR_ADD01
        .Tools(.Tools.Count - 1).Type = 0
        .Tools.Add ""
        .Tools(.Tools.Count - 1).Type = 2
        .Tools.Add TOOLBAR_ADD02
        .Tools(.Tools.Count - 1).Type = 0
        .Tools.Add ""
        .Tools(.Tools.Count - 1).Type = 2
        .Tools.Add TOOLBAR_ADD03
        .Tools(.Tools.Count - 1).Type = 0
        .Tools.Add ""
        .Tools(.Tools.Count - 1).Type = 2
    End With
    
End Sub

Private Sub Form_Resize()
           
    On Error Resume Next
           
    If Me.WindowState <> vbMinimized Then
        'フォームのサイズ変更にあわせて､ARViewerコントロールをリサイズします。
        With arv
            .Top = 0
            .Left = 0
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight
        End With
    End If

End Sub

Private Function OutPutPDF()
           
    Dim strFileName As String
           
    On Error GoTo OutPutPDF_Err
           
    '出力先指定
    CommonDialog1.FileName = m_objArPrint.Caption & ".pdf"
    CommonDialog1.DefaultExt = "pdf"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "pdfファイル (*.pdf)|*.pdf|すべてのファイル (*.*)|*.*"
    On Error Resume Next
    CommonDialog1.ShowSave
    If Err.Number = cdlCancel Then
        Exit Function
    End If
    On Error GoTo OutPutPDF_Err
    strFileName = CommonDialog1.FileName
    If strFileName = "" Then Exit Function
    
    Screen.MousePointer = vbHourglass

    Dim objARExport As New ARExportPDF
    Dim objARExportFix As New ARExportPDFFix
    
    objARExport.FileName = strFileName
    m_objArPrint.objReport.Run
    m_objArPrint.objReport.Export objARExport
    
    'PDFファイルの文字化けを修正
    objARExportFix.ExportFix objARExport.FileName

    Screen.MousePointer = vbDefault

    Exit Function

OutPutPDF_Err:

    Call MsgBox("PDF出力エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "OutPutPDF_Err")

End Function

