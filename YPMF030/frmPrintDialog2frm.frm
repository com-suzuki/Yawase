VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Begin VB.Form frmPrintDialog2 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "出品者伝票印刷"
   ClientHeight    =   2025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6510
   Icon            =   "frmPrintDialog2frm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Caption         =   "抽出条件"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4995
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'なし
         Height          =   615
         Left            =   1620
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   3315
         Begin VB.OptionButton optTaisyou 
            Caption         =   "出品者"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Top             =   120
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optTaisyou 
            Caption         =   "買　主"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   12
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.OptionButton optTanka 
         Caption         =   "売上単価"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3300
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optTanka 
         Caption         =   "仕入単価"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin imText6Ctl.imText txtOnum 
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   635
         Caption         =   "frmPrintDialog2frm.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog2frm.frx":007A
         Key             =   "frmPrintDialog2frm.frx":0098
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   4
         LengthAsByte    =   0
         Text            =   "9999"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   3
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   5
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "注文番号"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         TopMargin       =   0
         LabelTop        =   0
         LabelWidth      =   60
         LabelHeight     =   25
         LabelLeft       =   18
         BottomMargin    =   0
         RightMargin     =   0
         Spacing         =   0
         AutoAdjust      =   -1  'True
         BorderEffect    =   1
         BorderStyle     =   1
         LabelAutoSize   =   1
         LabelPosition   =   0
         ToolTip         =   ""
      End
      Begin imText6Ctl.imText txtOnum 
         Height          =   360
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   635
         Caption         =   "frmPrintDialog2frm.frx":00DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrintDialog2frm.frx":014A
         Key             =   "frmPrintDialog2frm.frx":0168
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   4
         LengthAsByte    =   0
         Text            =   "9999"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   3
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "印刷単価"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         TopMargin       =   0
         LabelTop        =   0
         LabelWidth      =   60
         LabelHeight     =   25
         LabelLeft       =   18
         BottomMargin    =   0
         RightMargin     =   0
         Spacing         =   0
         AutoAdjust      =   -1  'True
         BorderEffect    =   1
         BorderStyle     =   1
         LabelAutoSize   =   1
         LabelPosition   =   0
         ToolTip         =   ""
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "印刷対象"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         TopMargin       =   0
         LabelTop        =   0
         LabelWidth      =   60
         LabelHeight     =   25
         LabelLeft       =   18
         BottomMargin    =   0
         RightMargin     =   0
         Spacing         =   0
         AutoAdjust      =   -1  'True
         BorderEffect    =   1
         BorderStyle     =   1
         LabelAutoSize   =   1
         LabelPosition   =   0
         ToolTip         =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   420
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   660
      Width           =   1275
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmPrintDialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objArPrint As New clsArPrint
Public m_blnAutoPrint As Boolean

Private Sub cmdExecute_Click()
    
    On Error GoTo cmdExecute_Click_Err
    
    If optTaisyou(0).Value = True Then
        If PrintSyukkasya() = False Then Exit Sub
    ElseIf optTaisyou(1).Value = True Then
        If PrintKainusi() = False Then Exit Sub
    End If
    
    Exit Sub
    
cmdExecute_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("実行クリックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Form_KeyDown_Err

    'リターンキーで次のコントロールへフォーカス移動
    If KeyCode = vbKeyReturn And Shift <> vbShiftMask Then
        KeyCode = 0
        Global_SendKeys Me, VK_TAB
        Exit Sub
    End If
    
    Exit Sub

Form_KeyDown_Err:

    Call MsgBox("フォームキーダウン時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")
 
End Sub

Private Sub Form_Load()

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo Form_Load_Err

    txtOnum(0).Text = ""
    txtOnum(1).Text = ""
    
    With adoRecordset1
        strSQL = "SELECT Onum FROM DT030" & _
             " WHERE Odate = '" & g_strOdate & "'" & _
             " ORDER BY Onum"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            txtOnum(0).Text = .Fields("Onum")
            .MoveLast
            txtOnum(1).Text = .Fields("Onum")
        End If
        .Close
    End With
    
    Exit Sub
    
Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(txtOnum(0).Text) = "" Then
        strErrMsg = "注文番号を入力してください。"
        txtOnum(0).SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtOnum(1).Text) = "" Then
        strErrMsg = "注文番号を入力してください。"
        txtOnum(1).SetFocus
        GoTo ErrorTrap:
    End If
    
    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function

Private Sub optTaisyou_Click(Index As Integer)

    If optTaisyou(Index).Value = True Then
        optTanka(Index).Value = True
    End If

End Sub

Private Sub txtOnum_GotFocus(Index As Integer)

    txtOnum(Index).BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtOnum_LostFocus(Index As Integer)

    txtOnum(Index).BackColor = FOCUS_NO_COLOR
    
End Sub

Private Function MakeWork() As Boolean
    
    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim intLineCount As Integer
    Dim intIndex1 As Integer
    
    Dim curSkeep As Currency                '出品者維持管理費
    Dim curSrate As Currency                '出品者手数料率
    Dim intSfraction As Integer             '出品者端数処理
    Dim curSRounding As Currency            '出品者丸め単位
    Dim curTaxRate As Currency              '消費税率
    
    Dim curTotal As Currency
    Dim curCharge As Currency
    Dim curTax As Currency
    Dim curKeep As Currency
    Dim curGTotal As Currency
    
    On Error GoTo MakeWork_Err
    
    MakeWork = False
    
    Screen.MousePointer = vbHourglass
    
    '初期化
    curSkeep = 0        '出品者維持管理費
    intSfraction = 0    '出品者端数処理
    curSrate = 0        '出品者手数料率
    curSRounding = 0    '出品者丸め単位
    curTaxRate = 0      '消費税率
    
    '設定マスタオープン
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        If Not IsNull(adoMT010.Fields("Skeep")) Then curSkeep = adoMT010.Fields("Skeep")
        If Not IsNull(adoMT010.Fields("Sfraction")) Then intSfraction = adoMT010.Fields("Sfraction")
        If Not IsNull(adoMT010.Fields("Srate")) Then curSrate = adoMT010.Fields("Srate")
        If Not IsNull(adoMT010.Fields("SRounding")) Then curSRounding = adoMT010.Fields("SRounding")
    End If
    adoMT010.Close
    
    '消費税率取得
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, g_strOdate)
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF031"
    g_clsAdoAccess.Connection.Execute strSQL
    
    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF031"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
    'データオープン
    strSQL = "SELECT * FROM DT030" & _
             " WHERE Odate = '" & g_strOdate & "'" & _
             " AND Onum BETWEEN " & txtOnum(0).Text & " AND " & txtOnum(1).Text & _
             " ORDER BY Odate,Onum"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        curTotal = 0
        curCharge = 0
        curTax = 0
        curKeep = 0
        curGTotal = 0
        
        intLineCount = 1
        
        'データオープン
        strSQL = "SELECT * FROM DT031" & _
                 " WHERE Odate = '" & adoRecordset1.Fields("Odate") & "'" & _
                 " AND Onum = " & adoRecordset1.Fields("Onum") & _
                 " ORDER BY Odate,Onum,Line"
        adoRecordset2.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset2.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Key1") = adoRecordset1.Fields("Onum")
            wkRecordset.Fields("Key2") = adoRecordset1.Fields("Onum")
            wkRecordset.Fields("Num") = 1
            wkRecordset.Fields("Div") = "B"
            wkRecordset.Fields("Odate") = adoRecordset2.Fields("Odate")
            wkRecordset.Fields("Pnum") = adoRecordset1.Fields("Onum")
            wkRecordset.Fields("Scode") = adoRecordset1.Fields("Scode")
            wkRecordset.Fields("Sname") = Trim(adoRecordset1.Fields("Sname")) & "　様"
            wkRecordset.Fields("Line") = intLineCount
            wkRecordset.Fields("Icode") = adoRecordset2.Fields("Icode")
            wkRecordset.Fields("Iname") = adoRecordset2.Fields("Iname")
            wkRecordset.Fields("Qty") = adoRecordset2.Fields("Qty")
            If optTanka(0).Value = True Then
                '金額＝数量×仕入単価
                wkRecordset.Fields("Price1") = CCur(adoRecordset2.Fields("Qty")) * CCur(adoRecordset2.Fields("Price2"))
            ElseIf optTanka(1).Value = True Then
                '金額＝数量×売上単価
                wkRecordset.Fields("Price1") = CCur(adoRecordset2.Fields("Qty")) * CCur(adoRecordset2.Fields("Price1"))
            End If
            wkRecordset.Fields("Bcode") = adoRecordset2.Fields("Bcode")
'            wkRecordset.Fields("Bname") = adoRecordset2.Fields("Bname")
'            wkRecordset.Fields("Total") = 0
'            wkRecordset.Fields("Charge") = 0
'            wkRecordset.Fields("Tax") = 0
'            wkRecordset.Fields("Keep") = 0
'            wkRecordset.Fields("GTotal") = 0
            wkRecordset.Fields("Itime") = Format(Now(), "nn:ss")
            wkRecordset.Fields("Pcode") = g_strPcode
            wkRecordset.Fields("Pname") = g_strPname
            wkRecordset.Fields("Idiv") = 0
            wkRecordset.Fields("Result") = 0
            wkRecordset.Fields("Ocode") = Format(adoRecordset1.Fields("Onum"), "0000")
            wkRecordset.Fields("RePrint") = 0
            
            wkRecordset.Update
            
            curTotal = curTotal + CCur(wkRecordset.Fields("Price1"))
            
            adoRecordset2.MoveNext
            intLineCount = intLineCount + 1
        Loop
        adoRecordset2.Close
        
        '手数料
        If IsNull(adoRecordset1.Fields("ChargeDiv")) = False And adoRecordset1.Fields("ChargeDiv") = 1 Then
            curCharge = curTotal * curSrate / 100
            '切り捨て
            If IsNull(adoRecordset1.Fields("FixDiv")) = False And adoRecordset1.Fields("FixDiv") = 1 Then
                curCharge = Global_Rounding(curCharge, intSfraction, curSRounding)
            Else
                curCharge = Global_Rounding(curCharge, intSfraction, 1)
            End If
        End If
        '維持管理費
        If IsNull(adoRecordset1.Fields("KeepDiv")) = False And adoRecordset1.Fields("KeepDiv") = 1 Then
            curKeep = curSkeep
        End If
        '消費税計算
        If IsNull(adoRecordset1.Fields("TaxDiv")) = False And adoRecordset1.Fields("TaxDiv") = 1 Then
            '切り捨て
            If IsNull(adoRecordset1.Fields("FixDiv")) = False And adoRecordset1.Fields("FixDiv") = 1 Then
                '202308 インボイス対応
                curTax = Global_Get_Tax(curTotal - curCharge - curKeep, curTaxRate, intSfraction, 1)

            Else
                '202308 インボイス対応
                curTax = Global_Get_Tax(curTotal - curCharge - curKeep, curTaxRate, intSfraction, 1)
            End If
        End If
        
        '総合計＝合計－手数料－維持管理費
        curGTotal = curTotal - curCharge - curKeep + curTax
        'ワークの更新
        '202308 インボイス対応　出品者登録番号更新　WK_YPMF031.Invoice = '" & Trim(adoRecordset1.Fields("Addres")) & "'"
        strSQL = "UPDATE WK_YPMF031" & _
                 " SET WK_YPMF031.Total = " & curTotal & "," & _
                 " WK_YPMF031.Charge = " & curCharge & "," & _
                 " WK_YPMF031.Tax = " & curTax & "," & _
                 " WK_YPMF031.Keep = " & curKeep & "," & _
                 " WK_YPMF031.GTotal = " & curGTotal & "," & _
                 " WK_YPMF031.Bname = '" & Trim(adoRecordset1.Fields("Addres")) & "'" & _
                 " WHERE WK_YPMF031.Odate = '" & adoRecordset1.Fields("Odate") & "'" & _
                 " AND WK_YPMF031.Key1 = '" & adoRecordset1.Fields("Onum") & "'"

        g_clsAdoAccess.Connection.Execute strSQL


        
        adoRecordset1.MoveNext
    Loop
        
    adoRecordset1.Close
    wkRecordset.Requery     'バグ防止
    wkRecordset.Close
    
    Screen.MousePointer = vbDefault
    
    MakeWork = True
    
    Exit Function
    
MakeWork_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork_Err")
    
End Function

Private Function PrintSyukkasya() As Boolean
    
    Dim objRpt As New rptYpmf031
    
    On Error GoTo PrintSyukkasya_Err
    
    If DoValidationChecks() = False Then Exit Function
    If MakeWork() = False Then Exit Function
    
    If m_blnAutoPrint = False Then cmdCancel.SetFocus
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "支払伝票"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_Ypmf031"
        .Caption = "支払伝票"
        If m_blnAutoPrint = True Then
            If .PrintActiveReport(1) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        Else
            If .PrintActiveReport(0) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    End With
    
    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
PrintSyukkasya_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("出荷者印刷エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "PrintSyukkasya_Err")
    
End Function

Private Function PrintKainusi() As Boolean
    
    Dim objRpt As New rptYpmf032
    
    On Error GoTo PrintKainusi_Err
    
    If DoValidationChecks() = False Then Exit Function
    If MakeWork2() = False Then Exit Function
    
    If m_blnAutoPrint = False Then cmdCancel.SetFocus
    
    Screen.MousePointer = vbHourglass
    
    With objArPrint
        .Name = "買主伝票（注文）"
        .objReport = objRpt
        .Connection = g_clsAdoAccess.Connection
        .SQL = "SELECT * FROM QWK_Ypmf032"
        .Caption = "買主伝票（注文）"
        If m_blnAutoPrint = True Then
            If .PrintActiveReport(1) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        Else
            If .PrintActiveReport(0) = False Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    End With
    
    Set objRpt = Nothing
    Set objArPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
PrintKainusi_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("買主印刷エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "PrintKainusi_Err")
    
End Function

'目　的　　：印刷用ワーク作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０７／０８
'更新履歴　：
'
Private Function MakeWork2() As Boolean

    Dim strSQL As String
    Dim adoMT010 As New ADODB.Recordset
    Dim adoDT011 As New ADODB.Recordset
    Dim adoDT011M As New ADODB.Recordset
    Dim adoDT031 As New ADODB.Recordset
    Dim adoDT031M As New ADODB.Recordset
    Dim adoDT041 As New ADODB.Recordset
    Dim adoDT041TEMP As New ADODB.Recordset
    Dim wkRecordset As New ADODB.Recordset
    Dim wkRecordsetTemp As New ADODB.Recordset
    Dim strBuff1 As String
    Dim blnAddNewFlg As Boolean
    
    Dim intLine As Integer                  '行番号
    Dim lngCount As Long                    'レコード件数
    Dim curBkeep As Currency                '買主維持管理費(標準)
    Dim curBkeepCurrent As Currency         '買主維持管理費(今回)
    Dim intBfraction As Integer             '買主端数処理
    Dim intNum As Integer                   '回数
    Dim curTaxRate As Currency              '消費税率
    Dim intR As Integer                     '入金種別
    Dim curBRounding As Currency            '買主丸め単位
    Dim strMemo As String                   '伝票のメモ
    Dim strOdateNum As String

    Dim curPriceTotal As Currency
    Dim curTax As Currency
    Dim curGTotal As Currency

    On Error GoTo MakeWork2_Err
    
    MakeWork2 = False
    
    Screen.MousePointer = vbHourglass
    
    g_clsAdoSQL.Connection.BeginTrans
    
'********** 初期処理 **********
    
    '初期化
    lngCount = 0        'レコード件数
    curBkeep = 0        '買主維持管理費
    intBfraction = 0    '買主端数処理
    curTaxRate = 0      '消費税率
    intR = 0            '入金種別
    curBRounding = 0    '買主丸め単位
    strMemo = ""
    
    '設定マスタオープン
    strSQL = "{call sp_MT010;1}"
    adoMT010.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoMT010.EOF = False Then
        If Not IsNull(adoMT010.Fields("Bkeep")) Then curBkeep = adoMT010.Fields("Bkeep")
        If Not IsNull(adoMT010.Fields("Bfraction")) Then intBfraction = adoMT010.Fields("Bfraction")
        If Not IsNull(adoMT010.Fields("BRounding")) Then curBRounding = adoMT010.Fields("BRounding")
        If Not IsNull(adoMT010.Fields("Memo")) Then strMemo = adoMT010.Fields("Memo")
    End If
    adoMT010.Close
    
    '消費税率取得
    curTaxRate = Global_Get_TaxRate(g_clsAdoSQL, g_strOdate)
    
'********** ワーク **********
    
    'ワーク削除
    strSQL = "DELETE FROM WK_YPMF032"
    g_clsAdoAccess.Connection.Execute strSQL

    'ワークオープン
    strSQL = "SELECT * FROM WK_YPMF032"
    wkRecordset.Open strSQL, g_clsAdoAccess.Connection, adOpenKeyset, adLockOptimistic
    
'********** 注文データ **********
    
    '注文データオープン
    strSQL = "SELECT Bcode FROM DT031" & _
             " WHERE Bcode IS NOT NULL" & _
             " GROUP BY Bcode" & _
             " ORDER BY Bcode"
    adoDT031.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    
    lngCount = 0
    Do While Not adoDT031.EOF
        intLine = 1     '行番号
        curPriceTotal = 0
        
        '注文明細データオープン
        strSQL = "SELECT * FROM DT031" & _
                 " WHERE Bcode = " & adoDT031.Fields("Bcode") & _
                 " ORDER BY Bcode,Onum"
        adoDT031M.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoDT031M.EOF
            wkRecordset.AddNew
            wkRecordset.Fields("Key1") = Format(adoDT031M.Fields("Bcode"), "0000")
            wkRecordset.Fields("Key2") = Format(adoDT031M.Fields("Bcode"), "0000")
            wkRecordset.Fields("Num") = 1
            wkRecordset.Fields("Div") = "B"
            wkRecordset.Fields("Odate") = g_strOdate
            wkRecordset.Fields("Bcode") = adoDT031M.Fields("Bcode")
            wkRecordset.Fields("Line") = intLine
            wkRecordset.Fields("Bname") = Global_Get_Bname(g_clsAdoSQL, adoDT031M.Fields("Bcode"), g_strOdate, strBuff1)
            wkRecordset.Fields("Icode") = adoDT031M.Fields("Icode")
            wkRecordset.Fields("Iname") = adoDT031M.Fields("Iname")
            wkRecordset.Fields("Qty") = adoDT031M.Fields("Qty")
            If optTanka(0).Value = True Then
                wkRecordset.Fields("Price1") = CCur(adoDT031M.Fields("Qty")) * CCur(adoDT031M.Fields("Price2"))
            ElseIf optTanka(1).Value = True Then
                wkRecordset.Fields("Price1") = adoDT031M.Fields("Price")
            End If
            wkRecordset.Fields("Pnum") = adoDT031M.Fields("Onum")
            wkRecordset.Fields("Sname") = ""
'            wkRecordset.Fields("Total") = 0
'            wkRecordset.Fields("Tax") = 0
            wkRecordset.Fields("Keep") = curBkeep
'            wkRecordset.Fields("GTotal") = 0
            wkRecordset.Fields("Itime") = Format(Now(), "hh時mm分")
            wkRecordset.Fields("Pcode") = g_strPcode
            wkRecordset.Fields("Pname") = g_strPname
            wkRecordset.Fields("Idiv") = 0
            wkRecordset.Fields("Ocode") = right$(Format(adoDT031M.Fields("Onum"), "0000"), 4) & "*"
            wkRecordset.Fields("Memo") = strMemo
            wkRecordset.Update
        
            curPriceTotal = curPriceTotal + CCur(adoDT031M.Fields("Price"))
        
            adoDT031M.MoveNext
            intLine = intLine + 1
        Loop
        adoDT031M.Close
        
        '総合計計算
        curGTotal = curPriceTotal - curBkeep
        curTax = Global_Get_Tax(curPriceTotal, curTaxRate, intBfraction, curBRounding)
        
        
        'ワークの合計額などを更新
        strSQL = "UPDATE WK_YPMF032" & _
                 " SET WK_YPMF032.Total = " & curPriceTotal & "," & _
                 " WK_YPMF032.Tax = " & curTax & "," & _
                 " WK_YPMF032.Keep = " & curBkeep & "," & _
                 " WK_YPMF032.GTotal = " & curGTotal & _
                 " WHERE WK_YPMF032.Odate = '" & g_strOdate & "'" & _
                 " AND WK_YPMF032.Bcode = " & adoDT031.Fields("Bcode")
        g_clsAdoAccess.Connection.Execute strSQL
        
        adoDT031.MoveNext
        lngCount = lngCount + 1
    Loop
    adoDT031.Close
    
    wkRecordset.Requery
    wkRecordset.Close
    
    g_clsAdoSQL.Connection.CommitTrans
    
    If lngCount = 0 Then
        Call MsgBox("データがありません。", vbOKOnly + vbInformation, "")
        GoTo MakeWork2_Exit:
    End If
    
    MakeWork2 = True
    
MakeWork2_Exit:
    
    Screen.MousePointer = vbDefault
    
    Exit Function

MakeWork2_Cancel:

    g_clsAdoSQL.Connection.RollbackTrans
    GoTo MakeWork2_Exit:

MakeWork2_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    MakeWork2 = False
    Call MsgBox("印刷ワーク作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakeWork2_Err")
    GoTo MakeWork2_Exit:

End Function

