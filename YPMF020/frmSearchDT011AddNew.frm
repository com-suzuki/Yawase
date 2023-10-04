VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmSearchDT011AddNew 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "新規追加"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   10275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdExecute 
      Caption         =   "追加"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2100
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "キャンセル"
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
      Left            =   3300
      TabIndex        =   5
      Top             =   2100
      Width           =   1695
   End
   Begin CSComboLib.CSComboBox cboIcode 
      Height          =   360
      Left            =   6540
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1035
      _Version        =   262145
      _ExtentX        =   1826
      _ExtentY        =   635
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      ColDelim        =   ";"
      ColWidths       =   "2;20"
      Contents        =   "frmSearchDT011AddNew.frx":0000
      Extended        =   -1  'True
      ListBoxWidth    =   200
      MaxLength       =   5
      Text            =   "99999"
   End
   Begin imNumber6Ctl.imNumber imnQty 
      Height          =   435
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   767
      Calculator      =   "frmSearchDT011AddNew.frx":0019
      Caption         =   "frmSearchDT011AddNew.frx":0039
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011AddNew.frx":00A7
      Keys            =   "frmSearchDT011AddNew.frx":00C5
      Spin            =   "frmSearchDT011AddNew.frx":010F
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,##0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999
      MinValue        =   -999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2011496453
      Value           =   999999
      MaxValueVT      =   1230438405
      MinValueVT      =   1313734661
   End
   Begin imText6Ctl.imText txtIname 
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   8595
      _Version        =   65536
      _ExtentX        =   15161
      _ExtentY        =   714
      Caption         =   "frmSearchDT011AddNew.frx":0137
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011AddNew.frx":01A5
      Key             =   "frmSearchDT011AddNew.frx":01C3
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
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
      LengthAsByte    =   -1
      Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   1
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin imNumber6Ctl.imNumber imnNo 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   661
      Calculator      =   "frmSearchDT011AddNew.frx":0207
      Caption         =   "frmSearchDT011AddNew.frx":0227
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011AddNew.frx":0295
      Keys            =   "frmSearchDT011AddNew.frx":02B3
      Spin            =   "frmSearchDT011AddNew.frx":02FD
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#0"
      EditMode        =   0
      Enabled         =   0
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99
      MinValue        =   -99
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   2011496453
      Value           =   99
      MaxValueVT      =   1230438405
      MinValueVT      =   1313734661
   End
   Begin CSComboLib.CSComboBox cboIcode_Kana 
      Height          =   405
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      _Version        =   262145
      _ExtentX        =   8705
      _ExtentY        =   714
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColDelim        =   ";"
      ColWidths       =   "10;20;4"
      Contents        =   "frmSearchDT011AddNew.frx":0325
      Extended        =   -1  'True
      ListBoxWidth    =   700
      MaxLength       =   20
      Text            =   "WWWWWWWWWWWWWWWWWWWQ"
      ValueCol        =   2
   End
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
      _Version        =   262145
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "数　量"
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
      LabelWidth      =   40
      LabelHeight     =   25
      LabelLeft       =   28
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
      Index           =   9
      Left            =   60
      TabIndex        =   10
      Top             =   600
      Width           =   1455
      _Version        =   262145
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "植木ｶﾅ検索"
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
      LabelWidth      =   77
      LabelHeight     =   25
      LabelLeft       =   10
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
      Index           =   11
      Left            =   60
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
      _Version        =   262145
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "植木名"
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
      LabelWidth      =   45
      LabelHeight     =   25
      LabelLeft       =   26
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
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   10440
      TabIndex        =   0
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmSearchDT011AddNew.frx":033E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011AddNew.frx":03AC
      Key             =   "frmSearchDT011AddNew.frx":03CA
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   0
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   10
      LengthAsByte    =   -1
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin imText6Ctl.imText imtFocusEnd 
      Height          =   135
      Left            =   10620
      TabIndex        =   6
      Top             =   60
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmSearchDT011AddNew.frx":040E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011AddNew.frx":047C
      Key             =   "frmSearchDT011AddNew.frx":049A
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   0
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   10
      LengthAsByte    =   -1
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
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
      Index           =   3
      Left            =   60
      TabIndex        =   12
      Top             =   120
      Width           =   1455
      _Version        =   262145
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "行番号"
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
      LabelWidth      =   45
      LabelHeight     =   25
      LabelLeft       =   26
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
End
Attribute VB_Name = "frmSearchDT011AddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboIcode_Kana_Click()
    
    cboIcode_Kana.Tag = "1"
    cboIcode_Kana.BackColor = FOCUS_STOP_COLOR
    Call cboIcode_Kana_Validate(False)

End Sub

Private Sub cboIcode_Kana_DropDown()

    cboIcode_Kana.Tag = "1"
    Call MakecboIcode_Kana(cboIcode_Kana)

End Sub

Private Sub cboIcode_Kana_GotFocus()

    cboIcode_Kana.BackColor = FOCUS_STOP_COLOR
    cboIcode_Kana.Tag = ""
    Call SetImeMode(ActiveControl.hwnd, 9)
    
End Sub

Private Sub cboIcode_Kana_LostFocus()

    cboIcode_Kana.BackColor = FOCUS_NO_COLOR
    cboIcode_Kana.Tag = ""
    
End Sub

Private Sub cboIcode_Kana_Validate(Cancel As Boolean)

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo cboIcode_Kana_Validate_Err
    
    If Trim(cboIcode_Kana.Text) = "" Then Exit Sub
    If Trim(cboIcode_Kana.Tag) = "" Then Exit Sub
    If IsNumeric(cboIcode_Kana.Value) = False Then Exit Sub
    
    txtIname.Text = ""
    cboIcode.Text = cboIcode_Kana.Value
        
    '商品マスタ
    strSQL = "{call sp_MT050;2(" & cboIcode.Text & ")}"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = False Then
        txtIname.Text = IIf(IsNull(adoRecordset1.Fields("Iname")), "", Trim(adoRecordset1.Fields("Iname")))
    End If
    adoRecordset1.Close
    
    Exit Sub

cboIcode_Kana_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboIcode_Kana_Validate_Err")

End Sub

Private Sub MakecboIcode_Kana(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboIcode_Kana_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        '商品マスタ
        If Trim(strBuff1) = "" Then
            strSQL = "SELECT Ikana,Iname,Icode FROM MT050" & _
                     " ORDER BY Ikana,Iname,Icode"
        Else
            strSQL = "SELECT Ikana,Iname,Icode FROM MT050" & _
                     " WHERE Ikana LIKE '" & strBuff1 & "%'" & _
                     " ORDER BY Ikana,Iname,Icode"
        End If
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Ctrl.AddItem .Fields("Ikana") & ";" & .Fields("Iname") & ";" & .Fields("Icode")
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboIcode_Kana_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboIcode_Kana_Err")

End Sub

Private Sub cmdExecute_Click()

    Dim itmX As ListItem
    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo cmdExecute_Click_Err
    
    If MsgBox("追加しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    '入力チェック
    If DoValidationChecks() = False Then Exit Sub
    
    g_clsAdoSQL.Connection.BeginTrans
    
    With adoRecordset1
        strSQL = "SELECT * FROM DT011"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        .AddNew
        .Fields("Odate") = frmYpmf020.lblOdate.Caption
        .Fields("Pnum") = frmSearchDT011.imtPnum.Text
        .Fields("Line") = imnNo.Value
        .Fields("Icode") = IIf(IsNumeric(cboIcode.Text), cboIcode.Text, Null)
        .Fields("Iname") = txtIname.Text
        .Fields("Qty") = imnQty.Value
        .Fields("Idiv") = INPUT_OFF
        .Update
        .Close
    End With
    
    'リストビューにデータを追加
    Set itmX = frmSearchDT011.lsvMeisai.ListItems.Add(, , imnNo.Value, 0)
    itmX.SubItems(1) = Trim(cboIcode.Text)
    itmX.SubItems(2) = Trim(txtIname.Text)
    itmX.SubItems(3) = Format(imnQty.Value, "#,##0")

    If frmSearchDT011.lsvMeisai.Visible = False Then
        frmSearchDT011.lsvMeisai.Visible = True
        frmSearchDT011.lblDummy.Visible = False
    End If

    g_clsAdoSQL.Connection.CommitTrans

    Unload Me

    Exit Sub

cmdExecute_Click_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    Call MsgBox("追加クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

Private Sub cmdExit_Click()

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
    
    'ショートカットキーの割り当て
    Select Case KeyCode
        Case vbKeyF1
        Case vbKeyF2
        Case vbKeyF3
        Case vbKeyF4
        Case vbKeyF5
        Case vbKeyF6
        Case vbKeyF7
        Case vbKeyF8
        Case vbKeyF9
        Case vbKeyF10
        Case vbKeyF11
        Case vbKeyF12
        Case vbKeyF2
        Case vbKeyHome
        Case vbKeyPageUp
        Case vbKeyPageDown
    End Select

    Exit Sub

Form_KeyDown_Err:

    Call MsgBox("フォームキーダウン時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_KeyDown_Err")

End Sub

Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    cboIcode_Kana.Text = ""
    cboIcode_Kana.Tag = ""
    cboIcode.Text = ""
    txtIname.Text = ""
    imnQty.Value = 0
    imnNo.Value = frmSearchDT011.g_intAddNew_PnumLine

    Exit Sub
    
Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")

End Sub

Private Sub imnQty_GotFocus()
    
    imnQty.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnQty_LostFocus()
    
    imnQty.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus
    
End Sub

Private Sub imtFocusFirst_GotFocus()

    cboIcode_Kana.SetFocus

End Sub

Private Sub txtIname_GotFocus()

    txtIname.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtIname_LostFocus()

    txtIname.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err
    
    If Trim(txtIname.Text) = "" Then
        strErrMsg = "植木名を入力してください。"
        txtIname.SetFocus
        GoTo ErrorTrap:
    End If
    If imnQty.Value = 0 Then
        strErrMsg = "数量を入力してください。"
        imnQty.SetFocus
        GoTo ErrorTrap:
    End If

    DoValidationChecks = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks_Err:

    DoValidationChecks = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Err")

End Function
