VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmARPage 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ページ設定"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "frmARPage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   12
      Top             =   5100
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "O K"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   11
      Top             =   5100
      Width           =   1455
   End
   Begin VB.Frame fraMargin 
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   4395
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   6075
      Begin VB.TextBox txtRight 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   1620
         Width           =   915
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   2
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtBottom 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txtTop 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   0
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   1800
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   1800
         TabIndex        =   21
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   1800
         TabIndex        =   20
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1800
         TabIndex        =   19
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "右："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "左："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "下："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "上："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   420
         Width           =   435
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4875
      Left            =   60
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8599
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "余　白"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ページ"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraPage 
      BorderStyle     =   0  'なし
      Height          =   4335
      Left            =   120
      TabIndex        =   23
      Top             =   540
      Width           =   5955
      Begin VB.Frame Frame4 
         Caption         =   "給紙"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   60
         TabIndex        =   28
         Top             =   3420
         Width           =   4095
         Begin VB.ComboBox cboPaperSource 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmARPage.frx":000C
            Left            =   120
            List            =   "frmARPage.frx":0013
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   10
            Top             =   300
            Width           =   3855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "プリンタ"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   60
         TabIndex        =   27
         Top             =   1740
         Width           =   5835
         Begin VB.OptionButton optDeviceName2 
            Caption         =   "指定のプリンタ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   8
            Top             =   660
            Width           =   3315
         End
         Begin VB.OptionButton optDeviceName1 
            Caption         =   "通常使うプリンタ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   3315
         End
         Begin VB.ComboBox cboDeviceName 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   9
            Top             =   1140
            Width           =   5475
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "用紙"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   60
         TabIndex        =   25
         Top             =   840
         Width           =   5835
         Begin VB.ComboBox cboPaperSize 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1020
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   6
            Top             =   240
            Width           =   4515
         End
         Begin VB.Label Label1 
            Caption         =   "サイズ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   180
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "印刷の向き"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   2775
         Begin VB.OptionButton optOrientation2 
            Caption         =   "横"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1620
            TabIndex        =   5
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton optOrientation1 
            Caption         =   "縦"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   420
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "frmARPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_objArPrint As clsArPrint
Public m_blnCanselFlg As Boolean

'プリンタデバイスドライバの能力を取得する関数の宣言
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal pDevice As String, ByVal pPort As String, ByVal fwCapability As Long, pOutput As Any, pDevMode As Any) As Long
'ある位置から別の位置にメモリブロックを移動する関数の宣言
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'定数宣言
Private Const DC_PAPERS = 2
Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12      'プリンタで使用できる給紙方法の名前を取得することを示す
Private Const DC_PAPERNAMES = 16    'プリンタで使用できる用紙の名前を取得する定数

Private Sub cboDeviceName_DropDown()

    Dim objPrinter As Printer
    
    cboDeviceName.Clear
    For Each objPrinter In Printers
        'プリンタ名の取得
        cboDeviceName.AddItem objPrinter.DeviceName
    Next
    
End Sub

Private Sub cboPaperSize_DropDown()
    
    '用紙サイズの一覧表示
    Call ListPaperSizes

End Sub

Private Sub cboPaperSource_DropDown()
    
    Dim strBuff As String
    
    On Error GoTo cboPaperSource_DropDown_Err
    
    '給紙方法の一覧表示
    strBuff = cboPaperSource.Text
    Call ListPaperSource
    
    'コンボボックスの表示
    If GetCboListName(cboPaperSource, strBuff) = False Then
        cboPaperSource.Text = cboPaperSource.List(0)
    End If

    Exit Sub
    
cboPaperSource_DropDown_Err:
    
    Call MsgBox("給紙方法ドロップダウン時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "cboPaperSource_DropDown_Err")
    
End Sub

Private Sub cmdCancel_Click()

    m_blnCanselFlg = True
    Unload Me

End Sub

Private Sub cmdOk_Click()

    On Error GoTo cmdOk_Click_Err

    '入力チェック
    If IsNumeric(txtTop.Text) = False Then
        fraMargin.ZOrder 0
        txtTop.SetFocus
        Call MsgBox("正しい余白を入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    If IsNumeric(txtBottom.Text) = False Then
        fraMargin.ZOrder 0
        txtBottom.SetFocus
        Call MsgBox("正しい余白を入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    If IsNumeric(txtLeft.Text) = False Then
        fraMargin.ZOrder 0
        txtLeft.SetFocus
        Call MsgBox("正しい余白を入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    If IsNumeric(txtRight.Text) = False Then
        fraMargin.ZOrder 0
        txtRight.SetFocus
        Call MsgBox("正しい余白を入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    If Trim(cboPaperSize.Text) = "" Then
        fraPage.ZOrder 0
        cboPaperSize.SetFocus
        Call MsgBox("ページサイズを入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    If optDeviceName2.Value = True And Trim(cboDeviceName.Text) = "" Then
        fraPage.ZOrder 0
        cboPaperSize.SetFocus
        Call MsgBox("プリンタを入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If

    'プロパティセット
    m_objArPrint.PrnPageTopMargin = txtTop.Text
    m_objArPrint.PrnPageBottomMargin = txtBottom.Text
    m_objArPrint.PrnPageLeftMargin = txtLeft.Text
    m_objArPrint.PrnPageRightMargin = txtRight.Text
    
    If optOrientation1.Value = True Then
        m_objArPrint.PrnOrientation = vbPRORPortrait
    ElseIf optOrientation2.Value = True Then
        m_objArPrint.PrnOrientation = vbPRORLandscape
    End If
    m_objArPrint.PrnPaperSize = GetPaperSize()
    If optDeviceName1.Value = True Then
        m_objArPrint.PrnDefaultPrinter = True
        m_objArPrint.PrnDeviceName = ""
    ElseIf optDeviceName2.Value = True Then
        m_objArPrint.PrnDefaultPrinter = False
        m_objArPrint.PrnDeviceName = cboDeviceName.Text
    End If
    If cboPaperSource.ListIndex = 0 Then
        m_objArPrint.PrnPaperSource = 0 '自動の場合はゼロにしておく
    Else
        m_objArPrint.PrnPaperSource = cboPaperSource.ItemData(cboPaperSource.ListIndex)
    End If
    
    'フォームを閉じる
    Unload Me
    m_blnCanselFlg = False
    
    'プレビュー内容をリフレッシュ
    m_objArPrint.Refresh
    m_objArPrint.objReport.Restart
    m_objArPrint.objReport.Run
    frmARPreview.arv.ReportSource = m_objArPrint.objReport
    DoEvents
    
    Exit Sub
    
cmdOk_Click_Err:
    
    Call MsgBox("OKボタンクリック時エラー！！" _
                 & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdOk_Click_Err")
    
End Sub

Private Sub Form_Load()

    Dim intIndex1 As Integer

    On Error GoTo Form_Load_Err

    '初期設定
    fraMargin.ZOrder 0
    cboDeviceName.Enabled = False

    '初期表示
    txtTop.Text = m_objArPrint.PrnPageTopMargin
    txtBottom.Text = m_objArPrint.PrnPageBottomMargin
    txtLeft.Text = m_objArPrint.PrnPageLeftMargin
    txtRight.Text = m_objArPrint.PrnPageRightMargin
    
    Select Case m_objArPrint.PrnOrientation
        Case vbPRORPortrait:
            optOrientation1.Value = True
        Case vbPRORLandscape:
            optOrientation2.Value = True
        Case Else
            optOrientation2.Value = True
    End Select
    If m_objArPrint.PrnDefaultPrinter = True Then
        optDeviceName1.Value = True
    Else
        optDeviceName2.Value = True
        'コンボボックス作成
        Call cboDeviceName_DropDown
        For intIndex1 = 0 To cboDeviceName.ListCount - 1
            If Trim(cboDeviceName.List(intIndex1)) = Trim(m_objArPrint.PrnDeviceName) Then
                cboDeviceName.Text = cboDeviceName.List(intIndex1)
                Exit For
            End If
        Next intIndex1
    End If
    
    '用紙サイズ
    Call GetPaperSizeName(m_objArPrint.PrnPaperSize)
    
    '給紙方法
    Call GetPaperSource(m_objArPrint.PrnPaperSource)
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
 
End Sub

Private Sub optDeviceName1_Click()

    'プリンタ名のクリア
    cboDeviceName.Clear
    cboDeviceName.Enabled = False
    
End Sub

Private Sub optDeviceName2_Click()

    cboDeviceName.Enabled = True

End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.Index
        Case 1:
            fraMargin.ZOrder 0
        Case 2:
            fraPage.ZOrder 0
    End Select
    
End Sub

Private Function GetPaperSizeName(intPaperSize As Integer) As Boolean
    
    Dim strPrinterDeviceName     As String
    Dim strPrinterPortName       As String
    Dim lngDeviceCapability      As Long
    Dim lngSupportedPapersNeeded As Long
    Dim intSupportedPapers()     As Integer
    Dim lngSupportedPapersCount  As Long
    Dim lngWin32apiResultCode    As Long
    
    On Error GoTo GetPaperSizeName_Err
    
    If intPaperSize = 0 Then
        Exit Function
    End If
    
    'コンボボックス作成
    If ListPaperSizes() = False Then Exit Function
    
    'プリンタデバイス名とポート名を指定する
    If optDeviceName1.Value = True Then
        '通常使うプリンタの場合
        strPrinterDeviceName = Printer.DeviceName
        strPrinterPortName = Printer.Port
    ElseIf optDeviceName2.Value = True Then
        If Trim(cboDeviceName.Text) = "" Then Exit Function
        '指定のプリンタの場合
        Dim objPrinter As Printer
        
        For Each objPrinter In Printers
            If Trim(objPrinter.DeviceName) = Trim(cboDeviceName.Text) Then
                strPrinterDeviceName = objPrinter.DeviceName
                strPrinterPortName = objPrinter.Port
                Exit For
            End If
        Next
    End If
    
    '問い合わせる能力を指定
    lngDeviceCapability = DC_PAPERS
    'バッファの必要なサイズを取得
    lngSupportedPapersNeeded = DeviceCapabilities(strPrinterDeviceName, _
                                                  strPrinterPortName, _
                                                  lngDeviceCapability, _
                                                  ByVal vbNullString, _
                                                  ByVal vbNullString)
    'バッファを確保
    ReDim intSupportedPapers(lngSupportedPapersNeeded - 1)
    'サポートされている用紙サイズを取得
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, _
                                               strPrinterPortName, _
                                               lngDeviceCapability, _
                                               intSupportedPapers(0), _
                                               ByVal vbNullString)
    'バッファに用紙サイズが存在する間
    For lngSupportedPapersCount = 0 To lngSupportedPapersNeeded - 1
        'サポートされている用紙サイズを列挙
        If intSupportedPapers(lngSupportedPapersCount) = intPaperSize Then
            '用紙サイズ名のセット
            cboPaperSize.Text = cboPaperSize.List(lngSupportedPapersCount)
            Exit For
        End If
    Next lngSupportedPapersCount

    Exit Function
    
GetPaperSizeName_Err:

    GetPaperSizeName = ""
    Call MsgBox("用紙サイズ名取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "GetPaperSizeName_Err")
    
End Function

Private Function GetPaperSize() As Integer

    Dim strPrinterDeviceName     As String
    Dim strPrinterPortName       As String
    Dim lngDeviceCapability      As Long
    Dim lngSupportedPapersNeeded As Long
    Dim intSupportedPapers()     As Integer
    Dim lngSupportedPapersCount  As Long
    Dim lngWin32apiResultCode    As Long

    On Error GoTo GetPaperSize_Err
    
    GetPaperSize = 0
    
    'プリンタデバイス名とポート名を指定する
    If optDeviceName1.Value = True Then
        '通常使うプリンタの場合
        strPrinterDeviceName = Printer.DeviceName
        strPrinterPortName = Printer.Port
    ElseIf optDeviceName2.Value = True Then
        If Trim(cboDeviceName.Text) = "" Then Exit Function
        '指定のプリンタの場合
        Dim objPrinter As Printer
        
        For Each objPrinter In Printers
            If Trim(objPrinter.DeviceName) = Trim(cboDeviceName.Text) Then
                strPrinterDeviceName = objPrinter.DeviceName
                strPrinterPortName = objPrinter.Port
                Exit For
            End If
        Next
    End If
    
    '問い合わせる能力を指定
    lngDeviceCapability = DC_PAPERS
    'バッファの必要なサイズを取得
    lngSupportedPapersNeeded = DeviceCapabilities(strPrinterDeviceName, _
                                                  strPrinterPortName, _
                                                  lngDeviceCapability, _
                                                  ByVal vbNullString, _
                                                  ByVal vbNullString)
    'バッファを確保
    ReDim intSupportedPapers(lngSupportedPapersNeeded - 1)
    'サポートされている用紙サイズを取得
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, _
                                               strPrinterPortName, _
                                               lngDeviceCapability, _
                                               intSupportedPapers(0), _
                                               ByVal vbNullString)
    'バッファに用紙サイズが存在する間
    For lngSupportedPapersCount = 0 To lngSupportedPapersNeeded - 1
        If (cboPaperSize.ListCount - 1) < lngSupportedPapersCount Then Exit For
        If cboPaperSize.Text = cboPaperSize.List(lngSupportedPapersCount) Then
            GetPaperSize = intSupportedPapers(lngSupportedPapersCount)
            Exit For
        End If
    Next lngSupportedPapersCount

    Exit Function
    
GetPaperSize_Err:

    GetPaperSize = 0
    Call MsgBox("用紙サイズコード取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "GetPaperSize_Err")
    
End Function

Private Function ListPaperSizes() As Boolean
    
    Dim strPrinterDeviceName         As String
    Dim strPrinterPortName           As String
    Dim lngDeviceCapability          As Long
    Dim lngSupportedPaperNamesNeeded As Long
    Dim bytSupportedPaperNames()     As Byte
    Dim strSupportedPaperName        As String * 64
    Dim lngSupportedPaperNamesCount  As Long
    Dim lngWin32apiResultCode        As Long
    
    On Error GoTo ListPaperSizes_Err
    
    ListPaperSizes = False
    
    'コンボボックスを初期化
    cboPaperSize.Clear
    
    If optDeviceName1.Value = True Then
        '通常使うプリンタの場合
        'デバイス名とポート名を指定
        With Printer
            strPrinterDeviceName = .DeviceName
            strPrinterPortName = .Port
        End With
    ElseIf optDeviceName2.Value = True Then
        If Trim(cboDeviceName.Text) = "" Then Exit Function
        '指定のプリンタの場合
        Dim objPrinter As Printer
        
        For Each objPrinter In Printers
            If Trim(objPrinter.DeviceName) = Trim(cboDeviceName.Text) Then
                strPrinterDeviceName = objPrinter.DeviceName
                strPrinterPortName = objPrinter.Port
                Exit For
            End If
        Next
    End If
    
    '問い合わせる能力を指定
    lngDeviceCapability = DC_PAPERNAMES
    'バッファに必要なサイズを取得
    lngSupportedPaperNamesNeeded = DeviceCapabilities(strPrinterDeviceName, strPrinterPortName, lngDeviceCapability, ByVal vbNullString, ByVal vbNullString)
    'バッファを確保
    ReDim bytSupportedPaperNames(64 - 1, lngSupportedPaperNamesNeeded - 1)
    '使用できる用紙の名前を取得
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, strPrinterPortName, lngDeviceCapability, bytSupportedPaperNames(0, 0), ByVal vbNullString)
    '使用できる用紙の名前を列挙
    For lngSupportedPaperNamesCount = 0 To lngSupportedPaperNamesNeeded - 1
        '用紙の名前を複写
        MoveMemory ByVal strSupportedPaperName, bytSupportedPaperNames(0, lngSupportedPaperNamesCount), 64
        '用紙の名前を表示
        cboPaperSize.AddItem Left(strSupportedPaperName, InStr(strSupportedPaperName, vbNullChar) - 1)
    Next lngSupportedPaperNamesCount

    ListPaperSizes = True

    Exit Function

ListPaperSizes_Err:

    ListPaperSizes = False
    Call MsgBox("用紙サイズ一覧取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListPaperSizes_Err")
    
End Function

Private Function GetPaperSource(ByVal intPaperSource As Integer) As Boolean
    
    Dim intIndex1 As Integer
    
    On Error GoTo GetPaperSource_Err
    
    GetPaperSource = False
    
    If ListPaperSource() = False Then Exit Function
    
    With cboPaperSource
        .Text = .List(0)    '給紙方法「自動」を初期表示
        
        For intIndex1 = 0 To .ListCount - 1
            If .ItemData(intIndex1) = intPaperSource Then
                .Text = .List(intIndex1)
            End If
        Next intIndex1
    End With
            
    GetPaperSource = True

    Exit Function

GetPaperSource_Err:

    GetPaperSource = False
    Call MsgBox("給紙方法名取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "GetPaperSource_Err")
    
End Function

Private Function ListPaperSource() As Boolean
    
    Dim strPrinterDeviceName       As String
    Dim strPrinterPortName         As String
    Dim lngDeviceCapability        As Long
    Dim lngSupportedBinNamesNeeded As Long
    Dim bytSupportedBinNames()     As Byte
    Dim strSupportedBinName        As String * 24
    Dim lngSupportedBinNamesCount  As Long
    Dim lngWin32apiResultCode      As Long
    Dim intSupportedBins()         As Integer
    
    On Error GoTo ListPaperSource_Err
    
    ListPaperSource = False
    
    'コンボボックスを初期化
    cboPaperSource.Clear
    
    'デバイス名とポート名を指定
    If optDeviceName1.Value = True Then
        With Printer
            strPrinterDeviceName = .DeviceName
            strPrinterPortName = .Port
        End With
    Else
        If Trim$(cboDeviceName.Text) = "" Then Exit Function
        strPrinterDeviceName = cboDeviceName.Text
'        strPrinterPortName = lblPortName.Caption
    End If
    
    '問い合わせる能力を指定
    lngDeviceCapability = DC_BINNAMES
    
    'バッファに必要なサイズを取得
    lngSupportedBinNamesNeeded = DeviceCapabilities(strPrinterDeviceName, _
                                                    strPrinterPortName, _
                                                    lngDeviceCapability, _
                                                    ByVal vbNullString, _
                                                    ByVal vbNullString)
    'バッファを確保
    ReDim bytSupportedBinNames(24 - 1, lngSupportedBinNamesNeeded - 1)
    
    '使用できる給紙方法の名前を取得
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, _
                                               strPrinterPortName, _
                                               lngDeviceCapability, _
                                               bytSupportedBinNames(0, 0), _
                                               ByVal vbNullString)
    
    '問い合わせる能力を指定
    lngDeviceCapability = DC_BINS
    
    'バッファに必要なサイズを取得
    lngSupportedBinNamesNeeded = DeviceCapabilities(strPrinterDeviceName, _
                                                    strPrinterPortName, _
                                                    lngDeviceCapability, _
                                                    ByVal vbNullString, _
                                                    ByVal vbNullString)
    ' バッファを確保
    ReDim intSupportedBins(lngSupportedBinNamesNeeded - 1)
    
    '使用できる給紙方法を取得
    lngWin32apiResultCode = DeviceCapabilities(strPrinterDeviceName, _
                                               strPrinterPortName, _
                                               lngDeviceCapability, _
                                               intSupportedBins(0), _
                                               ByVal vbNullString)
    
    '使用できる給紙方法の名前を列挙
    With cboPaperSource
        'バッファに給紙方法の名前がある間
        For lngSupportedBinNamesCount = 0 To lngSupportedBinNamesNeeded - 1
            '給紙方法の名前を切り出し
            MoveMemory ByVal strSupportedBinName, bytSupportedBinNames _
                       (0, _
                       lngSupportedBinNamesCount), _
                        24
            
            '給紙方法の名前を表示
            .AddItem Left(strSupportedBinName, _
                     InStr(strSupportedBinName, _
                     vbNullChar) - 1)
            
            '給紙方法コードの設定
            .ItemData(lngSupportedBinNamesCount) = intSupportedBins(lngSupportedBinNamesCount)
        
        Next lngSupportedBinNamesCount
    End With
    
    ListPaperSource = True

    Exit Function

ListPaperSource_Err:

    ListPaperSource = False
    Call MsgBox("給紙方法一覧取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListPaperSource_Err")
    
End Function

Private Function GetCboListName(ByRef objCboControl As Control, ByVal varData As Variant) As Boolean
    
    Dim intIndex1 As Integer
    
    On Error GoTo GetCboListName_Err
    
    GetCboListName = False
    
    With objCboControl
        For intIndex1 = 0 To .ListCount - 1
            If CStr(.List(intIndex1)) = CStr(varData) Then
                .Text = .List(intIndex1)
                GetCboListName = True
                Exit Function
            End If
        Next intIndex1
    End With

    Exit Function

GetCboListName_Err:

    GetCboListName = False
    Call MsgBox("コンボボックスから名称取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "GetCboListName_Err")
    
End Function
