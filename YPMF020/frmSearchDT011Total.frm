VERSION 5.00
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "Cscomb32.ocx"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "Cscapt32.ocx"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Begin VB.Form frmSearchDT011Total 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "売立金額入力"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Ｏ　Ｋ"
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
      Left            =   1740
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   1695
      _Version        =   262145
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "合算売立金額"
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
      LabelWidth      =   90
      LabelHeight     =   25
      LabelLeft       =   11
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
   Begin imNumber6Ctl.imNumber imnPrice 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2595
      _Version        =   65536
      _ExtentX        =   4577
      _ExtentY        =   873
      Calculator      =   "frmSearchDT011Total.frx":0000
      Caption         =   "frmSearchDT011Total.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSearchDT011Total.frx":008E
      Keys            =   "frmSearchDT011Total.frx":00AC
      Spin            =   "frmSearchDT011Total.frx":00F6
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   -999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2011496453
      Value           =   999999999
      MaxValueVT      =   1230438405
      MinValueVT      =   1313734661
   End
   Begin CSComboLib.CSComboBox cboBcode 
      Height          =   480
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   1155
      _Version        =   262145
      _ExtentX        =   2037
      _ExtentY        =   847
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColDelim        =   ";"
      ColWidths       =   "4;20"
      Contents        =   "frmSearchDT011Total.frx":011E
      Extended        =   -1  'True
      ListBoxWidth    =   650
      MaxLength       =   4
      Text            =   "9999"
      ValueCol        =   0
   End
   Begin CSCaptLib.CSCaption csCaption1 
      Height          =   375
      Index           =   10
      Left            =   60
      TabIndex        =   5
      Top             =   720
      Width           =   1695
      _Version        =   262145
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "買主コード"
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
      LabelWidth      =   67
      LabelHeight     =   25
      LabelLeft       =   23
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
   Begin VB.Label lblBname 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Caption         =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWQ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3000
      TabIndex        =   6
      Top             =   720
      Width           =   8505
   End
End
Attribute VB_Name = "frmSearchDT011Total"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboBcode_Click()

    Call cboBcode_Validate(False)

End Sub

Private Sub cboBcode_DropDown()

    Call MakecboBcode(cboBcode)

End Sub

Private Sub cboBcode_GotFocus()
   
    cboBcode.BackColor = FOCUS_STOP_COLOR
    cboBcode.Tag = cboBcode.Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboBcode_LostFocus()
   
    cboBcode.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboBcode_Validate(Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cboBcode_Validate_Err
    
    If Trim(cboBcode.Text) = "" Then
        lblBname.Caption = ""
        Exit Sub
    End If
    If IsNumeric(cboBcode.Text) = False Then
        cboBcode.Text = ""
        lblBname.Caption = ""
        Exit Sub
    End If
    If cboBcode.Tag = cboBcode.Text Then Exit Sub
    
    lblBname.Caption = ""
    
    With adoRecordset1
        '得意先マスタ
        strSQL = "{call sp_MT070;2(" & Trim(cboBcode.Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If Not IsNull(.Fields("Fdiv")) Then
                If .Fields("Fdiv") = BUSINESS_DIV_BUYER Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    lblBname.Caption = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If lblBname.Caption = "" Then cboBcode.Text = ""
    
    Exit Sub

cboBcode_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboBcode_Validate_Err")

End Sub

Private Sub cmdExecute_Click()

    On Error GoTo cmdExecute_Click_Err

    If imnPrice.Value = 0 Then
        imnPrice.SetFocus
        DoEvents
        Call MsgBox("売立金額を入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If
    If Trim(cboBcode.Text) = "" Then
        cboBcode.SetFocus
        DoEvents
        Call MsgBox("買主コードを入力してください。", vbOKOnly + vbCritical, "入力チェック")
        Exit Sub
    End If

    frmSearchDT011.g_curPrice = imnPrice.Value
    frmSearchDT011.g_strBcode = Trim(cboBcode.Text)
    frmSearchDT011.g_strBname = Trim(lblBname.Caption)
    Unload Me

    Exit Sub

cmdExecute_Click_Err:

    Call MsgBox("ＯＫクリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdExecute_Click_Err")

End Sub

Private Sub cmdExit_Click()

    frmSearchDT011.g_curPrice = 0
    frmSearchDT011.g_strBcode = ""
    frmSearchDT011.g_strBname = ""
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

    imnPrice.Value = 0
    cboBcode.Text = ""
    cboBcode.Tag = ""
    lblBname.Caption = ""

End Sub

Private Sub imnPrice_GotFocus()
    
    imnPrice.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnPrice_LostFocus()
    
    imnPrice.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub MakecboBcode(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboBcode_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        strSQL = "SELECT * FROM vw_MT071" & _
                 " WHERE (Fdiv = " & BUSINESS_DIV_BUYER & " OR Fdiv = " & BUSINESS_DIV_ALL & ")"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            If Not IsNull(.Fields("Sdate")) And Not IsNull(.Fields("Fdate")) Then
                If .Fields("Sdate") <= Trim(frmYpmf020.lblOdate.Caption) And Trim(frmYpmf020.lblOdate.Caption) <= .Fields("Fdate") Then
                    Ctrl.AddItem .Fields("Bnum") & ";" & .Fields("Bname")
                Else
                    Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
                End If
            Else
                Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
            End If
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboBcode_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboBcode_Err")

End Sub


