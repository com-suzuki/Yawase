VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{E2690E23-9719-101B-9306-0020AF234C9D}#4.1#0"; "CSCMD32.OCX"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{93330F03-7CA6-101B-874B-0020AF109266}#4.1#0"; "CSCOMB32.OCX"
Object = "{CF720AD3-7E38-11CE-90BF-0000C037528B}#4.1#0"; "CSCAPT32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmYpmf030 
   BorderStyle     =   1  '固定(実線)
   ClientHeight    =   10485
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13095
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYpmf030.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   13095
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   11580
      TabIndex        =   75
      Top             =   0
      Width           =   1395
      Begin VB.CheckBox chkAutoCode 
         Caption         =   "ｺｰﾄﾞ自動採番"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   1935
      Left            =   60
      TabIndex        =   64
      Top             =   1260
      Width           =   12915
      Begin CSComboLib.CSComboBox txtSname 
         Height          =   405
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   7875
         _Version        =   262145
         _ExtentX        =   13891
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
         ColWidths       =   "0;40"
         Contents        =   "frmYpmf030.frx":0CFA
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   40
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         ValueCol        =   0
      End
      Begin VB.CommandButton cmdEtc 
         Caption         =   "..."
         Height          =   375
         Left            =   12480
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox chkKeepDiv 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9180
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkFixDiv 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11040
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CheckBox chkTaxDiv 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkChargeDiv 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdSearch2 
         Height          =   375
         Left            =   2340
         Picture         =   "frmYpmf030.frx":0D13
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CheckBox chkSoukin 
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1035
      End
      Begin imText6Ctl.imText txtAddres 
         Height          =   405
         Left            =   1560
         TabIndex        =   8
         Top             =   1440
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   714
         Caption         =   "frmYpmf030.frx":101D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":108B
         Key             =   "frmYpmf030.frx":10A9
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
         MultiLine       =   -1
         ScrollBars      =   2
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   80
         LengthAsByte    =   -1
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'なし
         Height          =   495
         Left            =   1500
         TabIndex        =   78
         Top             =   900
         Width           =   2175
         Begin VB.OptionButton optDiv 
            Caption         =   "市外"
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
            Left            =   1020
            TabIndex        =   6
            Top             =   120
            Width           =   915
         End
         Begin VB.OptionButton optDiv 
            Caption         =   "市内"
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
            TabIndex        =   5
            Top             =   120
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdMt070 
         Caption         =   "出品者のマスタ登録(F10)"
         Height          =   375
         Left            =   10020
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1020
         Visible         =   0   'False
         Width           =   2715
      End
      Begin CSComboLib.CSComboBox cboScode 
         Height          =   360
         Left            =   11580
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
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
         ColDelim        =   ";"
         ColWidths       =   "4;40"
         Contents        =   "frmYpmf030.frx":10ED
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   4
         Text            =   "9999"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   2
         Left            =   11400
         TabIndex        =   65
         Top             =   180
         Visible         =   0   'False
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "出品者検索"
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
         LabelWidth      =   75
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   4
         Left            =   60
         TabIndex        =   66
         Top             =   1440
         Width           =   1485
         _Version        =   262145
         _ExtentX        =   2619
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "登録番号"
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
         LabelLeft       =   19
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
         Index           =   6
         Left            =   9180
         TabIndex        =   67
         Top             =   1440
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "電話番号"
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
      Begin imText6Ctl.imText txtTel 
         Height          =   360
         Left            =   10680
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   635
         Caption         =   "frmYpmf030.frx":1106
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":1174
         Key             =   "frmYpmf030.frx":1192
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
         MaxLength       =   15
         LengthAsByte    =   -1
         Text            =   "999999999999999"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin CSComboLib.CSComboBox cboScode_Kana 
         Height          =   360
         Left            =   10020
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   4875
         _Version        =   262145
         _ExtentX        =   8599
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
         ColDelim        =   ";"
         ColWidths       =   "10;20;4"
         Contents        =   "frmYpmf030.frx":11C6
         Extended        =   -1  'True
         ListBoxWidth    =   600
         MaxLength       =   20
         Text            =   "WWWWWWWWWWWWWWWWWWWQ"
         ValueCol        =   2
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   10
         Left            =   60
         TabIndex        =   68
         Top             =   600
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "出品者名"
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
         Index           =   12
         Left            =   60
         TabIndex        =   77
         Top             =   1020
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "地区区分"
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
         Index           =   22
         Left            =   4200
         TabIndex        =   95
         Top             =   1020
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "送　金"
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
         Index           =   5
         Left            =   60
         TabIndex        =   96
         Top             =   180
         Visible         =   0   'False
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "受付番号"
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
      Begin imText6Ctl.imText txtPnum 
         Height          =   360
         Left            =   1560
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   635
         Caption         =   "frmYpmf030.frx":11DF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":124D
         Key             =   "frmYpmf030.frx":126B
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
      Begin CSCaptLib.CSCaption lblFixDiv 
         Height          =   375
         Left            =   9660
         TabIndex        =   102
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "切り捨て"
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
         LabelWidth      =   55
         LabelHeight     =   25
         LabelLeft       =   15
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
      Begin CSCaptLib.CSCaption lblTaxDiv 
         Height          =   375
         Left            =   5940
         TabIndex        =   103
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "消費税"
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
         LabelLeft       =   20
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
      Begin CSCaptLib.CSCaption lblChargeDiv 
         Height          =   375
         Left            =   4140
         TabIndex        =   104
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "手数料"
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
         LabelLeft       =   20
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
      Begin CSCaptLib.CSCaption lblKeepDiv 
         Height          =   375
         Left            =   7800
         TabIndex        =   106
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "維持管理費"
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
         LabelWidth      =   75
         LabelHeight     =   25
         LabelLeft       =   5
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
         Caption         =   "↓出品者名の後に自動的に(注文分)がつきます"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   109
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.Frame fraLogin 
      Height          =   675
      Left            =   3120
      TabIndex        =   55
      Top             =   600
      Width           =   9855
      Begin VB.CommandButton cmdLogin 
         Caption         =   "開催年月日と担当者の変更"
         Height          =   375
         Left            =   6960
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   180
         Width           =   2655
      End
      Begin CSComboLib.CSComboBox cboPcode 
         Height          =   360
         Left            =   9900
         TabIndex        =   56
         Top             =   180
         Visible         =   0   'False
         Width           =   675
         _Version        =   262145
         _ExtentX        =   1191
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
         ColDelim        =   ";"
         ColWidths       =   "2;20"
         Contents        =   "frmYpmf030.frx":12AF
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   2
         Text            =   "99"
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   59
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "開催年月日"
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
         LabelWidth      =   75
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   8
         Left            =   3480
         TabIndex        =   60
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "担当者"
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
      Begin VB.Label lblPname 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "ＮＮＮＮＮＮＮＮＮＮ"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   4980
         TabIndex        =   58
         Top             =   180
         Width           =   1905
      End
      Begin VB.Label lblOdate 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "9999/12/31"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1620
         TabIndex        =   57
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame fraRecordSelector 
      Height          =   615
      Left            =   9240
      TabIndex        =   47
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmdDataMove 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   60
         Picture         =   "frmYpmf030.frx":12C8
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   600
         Picture         =   "frmYpmf030.frx":1412
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1140
         Picture         =   "frmYpmf030.frx":155C
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
      Begin VB.CommandButton cmdDataMove 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1680
         Picture         =   "frmYpmf030.frx":16A6
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   180
         Width           =   550
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   735
      Left            =   60
      TabIndex        =   46
      Top             =   9720
      Width           =   12915
      Begin CSCmdLibCtl.CSCmdBtn cmdClear 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "画面ｸﾘｱ(F8)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   2
         rPic.top        =   4
         rPic.right      =   0
         rPic.bottom     =   0
         rText.left      =   10
         rText.top       =   8
         rText.right     =   103
         rText.bottom    =   27
         Picture         =   "frmYpmf030.frx":17F0
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExit 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   11100
         TabIndex        =   36
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "終了(F9)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SizePicture     =   -1  'True
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   14
         rPic.top        =   6
         rPic.right      =   21
         rPic.bottom     =   21
         rText.left      =   43
         rText.top       =   8
         rText.right     =   109
         rText.bottom    =   27
         Picture         =   "frmYpmf030.frx":180C
      End
      Begin CSCmdLibCtl.CSCmdBtn cmdExecute 
         Height          =   495
         Left            =   9360
         TabIndex        =   34
         Top             =   180
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "実行(F12)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SizePicture     =   -1  'True
         OwnPicDisabled  =   0   'False
         CaptionPosition =   3
         rPic.left       =   9
         rPic.top        =   6
         rPic.right      =   21
         rPic.bottom     =   21
         rText.left      =   34
         rText.top       =   8
         rText.right     =   109
         rText.bottom    =   27
         Picture         =   "frmYpmf030.frx":1966
      End
   End
   Begin VB.Frame fraSyori 
      Height          =   615
      Left            =   60
      TabIndex        =   40
      Top             =   0
      Width           =   9135
      Begin VB.OptionButton optSyori 
         Caption         =   "集計表印刷"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7740
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   180
         Width           =   1335
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "伝票印刷"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6540
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "確認表印刷"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5160
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   180
         Width           =   1395
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "削　除"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3960
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "変　更"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2760
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optSyori 
         Caption         =   "新　規"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1560
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   44
         Top             =   180
         Width           =   1455
         _Version        =   262145
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "処理区分"
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
   End
   Begin VB.Frame fraMeisai 
      Height          =   6555
      Left            =   60
      TabIndex        =   39
      Top             =   3180
      Width           =   12915
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   2340
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmYpmf030.frx":1DB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "明細挿入(&I)"
         Height          =   375
         Left            =   9300
         TabIndex        =   31
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdPast 
         Caption         =   "明細貼付(&P)"
         Height          =   375
         Left            =   7740
         TabIndex        =   30
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CheckBox chkChumon 
         Caption         =   "注 文 分(&O)"
         Height          =   435
         Left            =   3060
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Frame fraDetail 
         Height          =   2355
         Left            =   9840
         TabIndex        =   73
         Top             =   120
         Visible         =   0   'False
         Width           =   2955
         Begin VB.CheckBox chkSdiv 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   86
            Top             =   180
            Width           =   615
         End
         Begin VB.CheckBox chkBdiv 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   85
            Top             =   600
            Width           =   615
         End
         Begin VB.CheckBox chkIdiv 
            Height          =   435
            Left            =   2280
            TabIndex        =   74
            Top             =   1860
            Width           =   435
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   87
            Top             =   180
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "出品伝票出力済み"
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
            LabelWidth      =   120
            LabelHeight     =   25
            LabelLeft       =   8
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
            Index           =   17
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "買主伝票出力済み"
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
            LabelWidth      =   120
            LabelHeight     =   25
            LabelLeft       =   8
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
            Index           =   18
            Left            =   120
            TabIndex        =   89
            Top             =   1020
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "買主精算回数"
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
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   90
            Top             =   1440
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "出品者精算回数"
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
            LabelWidth      =   105
            LabelHeight     =   25
            LabelLeft       =   16
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
         Begin imNumber6Ctl.imNumber imnBnum 
            Height          =   375
            Left            =   2220
            TabIndex        =   91
            Top             =   1020
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   661
            Calculator      =   "frmYpmf030.frx":1F12
            Caption         =   "frmYpmf030.frx":1F32
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf030.frx":1FA0
            Keys            =   "frmYpmf030.frx":1FBE
            Spin            =   "frmYpmf030.frx":2008
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
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#0"
            HighlightText   =   1
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
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin imNumber6Ctl.imNumber imnSnum 
            Height          =   375
            Left            =   2220
            TabIndex        =   92
            Top             =   1440
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   661
            Calculator      =   "frmYpmf030.frx":2030
            Caption         =   "frmYpmf030.frx":2050
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf030.frx":20BE
            Keys            =   "frmYpmf030.frx":20DC
            Spin            =   "frmYpmf030.frx":2126
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
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#0"
            HighlightText   =   1
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
            ShowContextMenu =   1
            ValueVT         =   2012217349
            Value           =   99
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   93
            Top             =   1860
            Width           =   2055
            _Version        =   262145
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "入力済みフラグ"
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
            LabelWidth      =   97
            LabelHeight     =   25
            LabelLeft       =   20
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
      Begin VB.CommandButton cmdDetail 
         Caption         =   "..."
         Height          =   375
         Left            =   2040
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "明細コピー(&C)"
         Height          =   375
         Left            =   6180
         TabIndex        =   29
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear_Dst 
         Caption         =   "明細クリア(&N)"
         Height          =   375
         Left            =   4620
         TabIndex        =   28
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdMt050 
         Caption         =   "植木のマスタ登録(F11)"
         Height          =   375
         Left            =   6660
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   600
         Width           =   2415
      End
      Begin CSComboLib.CSComboBox cboIcode 
         Height          =   360
         Left            =   9000
         TabIndex        =   15
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
         Contents        =   "frmYpmf030.frx":214E
         Extended        =   -1  'True
         ListBoxWidth    =   200
         MaxLength       =   5
         Text            =   "99999"
      End
      Begin MSComctlLib.ListView lsvMeisai 
         Height          =   3135
         Left            =   120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2940
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "行"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "コード"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "植木名称"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "数　量"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "済"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "仕入単価"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "単　価"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "売立金額"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "買主"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "買主名称"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "出品者伝票区分"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "買主伝票区分"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "買主精算回数"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "出品者精算回数"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "明細削除(&D)"
         Height          =   375
         Left            =   3060
         TabIndex        =   27
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "明細登録(&A)"
         Height          =   375
         Left            =   1500
         TabIndex        =   26
         Top             =   2520
         Width           =   1575
      End
      Begin imNumber6Ctl.imNumber imnQty 
         Height          =   435
         Left            =   1560
         TabIndex        =   20
         Top             =   1560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   767
         Calculator      =   "frmYpmf030.frx":2167
         Caption         =   "frmYpmf030.frx":2187
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":21F5
         Keys            =   "frmYpmf030.frx":2213
         Spin            =   "frmYpmf030.frx":225D
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
         ValueVT         =   2011365381
         Value           =   999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin imText6Ctl.imText txtIname 
         Height          =   405
         Left            =   1560
         TabIndex        =   19
         Top             =   1080
         Width           =   8595
         _Version        =   65536
         _ExtentX        =   15161
         _ExtentY        =   714
         Caption         =   "frmYpmf030.frx":2285
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":22F3
         Key             =   "frmYpmf030.frx":2311
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   3
         Left            =   60
         TabIndex        =   45
         Top             =   180
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
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   54
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
      Begin imNumber6Ctl.imNumber imnNo 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   180
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   661
         Calculator      =   "frmYpmf030.frx":2355
         Caption         =   "frmYpmf030.frx":2375
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":23E3
         Keys            =   "frmYpmf030.frx":2401
         Spin            =   "frmYpmf030.frx":244B
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
         ValueVT         =   2011365381
         Value           =   99
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSComboLib.CSComboBox cboIcode_Kana 
         Height          =   405
         Left            =   1560
         TabIndex        =   16
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
         Contents        =   "frmYpmf030.frx":2473
         Extended        =   -1  'True
         ListBoxWidth    =   700
         MaxLength       =   20
         Text            =   "WWWWWWWWWWWWWWWWWWWQ"
         ValueCol        =   2
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   9
         Left            =   60
         TabIndex        =   62
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
      Begin imNumber6Ctl.imNumber imnQty_Total 
         Height          =   375
         Left            =   4920
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   6120
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   661
         Calculator      =   "frmYpmf030.frx":248C
         Caption         =   "frmYpmf030.frx":24AC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":251A
         Keys            =   "frmYpmf030.frx":2538
         Spin            =   "frmYpmf030.frx":257A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999
         MinValue        =   -999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2012217349
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   11
         Left            =   60
         TabIndex        =   69
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
      Begin VB.Frame fraChumon 
         BorderStyle     =   0  'なし
         Height          =   1095
         Left            =   4440
         TabIndex        =   79
         Top             =   1380
         Width           =   8415
         Begin imNumber6Ctl.imNumber imnPrice1 
            Height          =   435
            Left            =   3900
            TabIndex        =   23
            Top             =   180
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   767
            Calculator      =   "frmYpmf030.frx":25A2
            Caption         =   "frmYpmf030.frx":25C2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf030.frx":2630
            Keys            =   "frmYpmf030.frx":264E
            Spin            =   "frmYpmf030.frx":2698
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   9999999
            MinValue        =   -9999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2088828933
            Value           =   9999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   13
            Left            =   2880
            TabIndex        =   80
            Top             =   180
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "単　価"
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
            LabelLeft       =   12
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
            Height          =   435
            Left            =   6720
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   180
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   767
            Calculator      =   "frmYpmf030.frx":26C0
            Caption         =   "frmYpmf030.frx":26E0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf030.frx":274E
            Keys            =   "frmYpmf030.frx":276C
            Spin            =   "frmYpmf030.frx":27B6
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
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
            ValueVT         =   2012217349
            Value           =   999999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   14
            Left            =   5460
            TabIndex        =   81
            Top             =   180
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "売立金額"
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
         Begin CSComboLib.CSComboBox cboBcode 
            Height          =   435
            Left            =   1080
            TabIndex        =   25
            Top             =   660
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
            _ExtentY        =   767
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColDelim        =   ";"
            ColWidths       =   "4;20"
            Contents        =   "frmYpmf030.frx":27DE
            Extended        =   -1  'True
            ListBoxWidth    =   500
            MaxLength       =   4
            Text            =   "9999"
            ValueCol        =   0
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   15
            Left            =   60
            TabIndex        =   82
            Top             =   660
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "買　主"
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
            LabelLeft       =   12
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
         Begin imNumber6Ctl.imNumber imnPrice2 
            Height          =   435
            Left            =   1260
            TabIndex        =   22
            Top             =   180
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   767
            Calculator      =   "frmYpmf030.frx":27F7
            Caption         =   "frmYpmf030.frx":2817
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmYpmf030.frx":2885
            Keys            =   "frmYpmf030.frx":28A3
            Spin            =   "frmYpmf030.frx":28ED
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   9999999
            MinValue        =   -9999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2088828933
            Value           =   9999999
            MaxValueVT      =   1230438405
            MinValueVT      =   1313734661
         End
         Begin CSCaptLib.CSCaption csCaption1 
            Height          =   375
            Index           =   21
            Left            =   60
            TabIndex        =   94
            Top             =   180
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "仕入単価"
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
            LabelLeft       =   8
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
            Height          =   405
            Left            =   2160
            TabIndex        =   83
            Top             =   660
            Width           =   6045
         End
      End
      Begin imNumber6Ctl.imNumber imnPrice_Total 
         Height          =   375
         Left            =   8940
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   6120
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         Calculator      =   "frmYpmf030.frx":2915
         Caption         =   "frmYpmf030.frx":2935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":29A3
         Keys            =   "frmYpmf030.frx":29C1
         Spin            =   "frmYpmf030.frx":2A03
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999
         MinValue        =   -999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2011496453
         Value           =   999999999999
         MaxValueVT      =   1230438405
         MinValueVT      =   1313734661
      End
      Begin VB.Label Label2 
         Caption         =   "合　計"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3780
         TabIndex        =   63
         Top             =   6180
         Width           =   1095
      End
   End
   Begin VB.Frame fraKey 
      Height          =   675
      Left            =   60
      TabIndex        =   38
      Top             =   600
      Width           =   3015
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   2340
         Picture         =   "frmYpmf030.frx":2A2B
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   555
      End
      Begin CSCaptLib.CSCaption csCaption1 
         Height          =   375
         Index           =   23
         Left            =   60
         TabIndex        =   97
         Top             =   180
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
         Left            =   1560
         TabIndex        =   1
         Top             =   180
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   635
         Caption         =   "frmYpmf030.frx":2D35
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmYpmf030.frx":2DA3
         Key             =   "frmYpmf030.frx":2DC1
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
   End
   Begin imText6Ctl.imText imtFocusFirst 
      Height          =   135
      Left            =   13200
      TabIndex        =   0
      Top             =   180
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf030.frx":2E05
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf030.frx":2E73
      Key             =   "frmYpmf030.frx":2E91
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
      Left            =   13380
      TabIndex        =   37
      Top             =   180
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf030.frx":2ED5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf030.frx":2F43
      Key             =   "frmYpmf030.frx":2F61
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
   Begin imText6Ctl.imText imtScode_Kana_Focus1 
      Height          =   135
      Left            =   13200
      TabIndex        =   11
      Top             =   480
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf030.frx":2FA5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf030.frx":3013
      Key             =   "frmYpmf030.frx":3031
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
   Begin imText6Ctl.imText imtScode_Kana_Focus2 
      Height          =   135
      Left            =   13380
      TabIndex        =   12
      Top             =   480
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf030.frx":3075
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf030.frx":30E3
      Key             =   "frmYpmf030.frx":3101
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
   Begin imText6Ctl.imText imtIcode_Kana_Focus1 
      Height          =   135
      Left            =   13200
      TabIndex        =   17
      Top             =   3660
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf030.frx":3145
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf030.frx":31B3
      Key             =   "frmYpmf030.frx":31D1
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
   Begin imText6Ctl.imText imtIcode_Kana_Focus2 
      Height          =   135
      Left            =   13380
      TabIndex        =   18
      Top             =   3660
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   238
      Caption         =   "frmYpmf030.frx":3215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYpmf030.frx":3283
      Key             =   "frmYpmf030.frx":32A1
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
End
Attribute VB_Name = "frmYpmf030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_strLastOnum As String                      '最後に登録した受付番号

Const AUTO_CODE = 1                                 'コードの自動採番
Const MAX_ROW = 20                                  '明細の最大行数
Const MAX_COL = 13                                  '明細の列数
Const DETAIL_FORECOLOR1 = &HFF&
Const DETAIL_FORECOLOR2 = &HFF0000

Private Type typDetail
    Div     As Boolean
    Field01 As Variant
    Field02 As Variant
    Field03 As Variant
    Field04 As Variant
    Field05 As Variant
    Field06 As Variant
    Field07 As Variant
    Field08 As Variant
    Field09 As Variant
    Field10 As Variant
    Field11 As Variant
    Field12 As Variant
    Field13 As Variant
End Type
Private m_typDetailCopy As typDetail                '明細のコピー/貼り付け用

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

Private Sub cboScode_DropDown()

    Call MakecboScode(cboScode)
    
End Sub

Private Sub cboScode_GotFocus()

    cboScode.BackColor = FOCUS_STOP_COLOR
    cboScode.Tag = cboScode.Text
    Call SetImeMode(ActiveControl.hwnd, 2)
    
End Sub

Private Sub cboScode_LostFocus()

    cboScode.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub cboScode_Kana_Click()
    
    cboScode_Kana.Tag = "1"
    cboScode_Kana.BackColor = FOCUS_STOP_COLOR
    Call cboScode_Kana_Validate(False)
    
End Sub

Private Sub cboScode_Kana_DropDown()

    cboScode_Kana.Tag = "1"
    Call MakecboScode_Kana(cboScode_Kana)

End Sub

Private Sub cboScode_Kana_GotFocus()

    cboScode_Kana.BackColor = FOCUS_STOP_COLOR
    cboScode_Kana.Tag = ""
    Call SetImeMode(ActiveControl.hwnd, 9)

End Sub

Private Sub cboScode_Kana_LostFocus()

    cboScode_Kana.BackColor = FOCUS_NO_COLOR
    cboScode_Kana.Tag = ""

End Sub

Private Sub cboScode_Kana_Validate(Cancel As Boolean)

    On Error GoTo cboScode_Kana_Validate_Err
    
    If Trim(cboScode_Kana.Tag) = "" Then Exit Sub
    If Trim(cboScode_Kana.Tag) = "" Then Exit Sub
    If IsNumeric(cboScode_Kana.Value) = False Then Exit Sub
    
    cboScode.Text = cboScode_Kana.Value
    Call cboScode_Validate(False)
    
    Exit Sub

cboScode_Kana_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboScode_Kana_Validate_Err")

End Sub

Private Sub cboScode_Validate(Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim strBuff As String

    On Error GoTo cboScod_Validate_Err
    
    If Trim(cboScode.Text) = "" Then Exit Sub
    If IsNumeric(cboScode.Text) = False Then
        cboScode.Text = ""
        Exit Sub
    End If
    If cboScode.Tag = cboScode.Text Then Exit Sub
    
    txtSname.Text = ""
    txtAddres.Text = ""
    txtTel.Text = ""
    
    With adoRecordset1
        '得意先マスタ
        strSQL = "{call sp_MT070;2(" & Trim(cboScode.Text) & ")}"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            If IsNull(.Fields("Fdiv")) = False Then
                If .Fields("Fdiv") = BUSINESS_DIV_EXHIBITION Or .Fields("Fdiv") = BUSINESS_DIV_ALL Then
                    txtSname.Text = IIf(IsNull(.Fields("Bname")), "", Trim(.Fields("Bname")))
                    txtAddres.Text = IIf(IsNull(.Fields("Addres")), "", Trim(.Fields("Addres")))
                    txtTel.Text = IIf(IsNull(.Fields("Tel")), "", Trim(.Fields("Tel")))
                End If
            End If
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    If Trim(txtSname.Text) = "" Then cboScode.Text = ""
    
    Exit Sub

cboScod_Validate_Err:

    Call MsgBox("フォーカス移動前エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "cboScod_Validate_Err")

End Sub

'目　的　　：
'条　件　　：ｺｰﾄﾞ自動採番クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub chkAutoCode_Click()

    On Error Resume Next

    If optSyori(0).Value = True And chkAutoCode.Value = 1 Then
         txtOnum.Text = AutoCodeSet()
         If txtOnum.Enabled Then txtOnum.SetFocus
    End If

End Sub

Private Sub chkChumon_Click()

    On Error GoTo chkChumon_Click_Err

    If chkChumon.Value = 1 Then
        chkChumon.BackColor = BUTTON_ON
        fraChumon.Visible = True
    Else
        chkChumon.BackColor = BUTTON_OFF
        fraChumon.Visible = False
    End If
    
    DoEvents

    Exit Sub

chkChumon_Click_Err:

    Call MsgBox("注文分クリック時エラー！！" _
                    & vbCrLf & Error$, vbOKOnly + vbCritical, "chkChumon_Click_Err")

End Sub

'目　的　　：
'条　件　　：画面クリアクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub cmdClear_Click()

    Call FieldsClear(0)
    On Error Resume Next
    If optSyori(0).Value = True Then
        If chkAutoCode.Value = 1 Then txtOnum.Text = AutoCodeSet()
    End If
    txtOnum.SetFocus

End Sub

Private Sub cmdClear_Dst_Click()

    Call FieldsClear(3)
    Call ListViewGetMaxRow
    cboIcode_Kana.SetFocus

End Sub

Private Sub cmdCopy_Click()

    On Error GoTo cmdCopy_Click_Err

    Call ListViewGetMaxRow
    '明細クリア
    chkIdiv.Value = 0
    chkSdiv.Value = 0
    chkBdiv.Value = 0
    imnBnum.Value = 0
    imnSnum.Value = 0
    
    m_typDetailCopy.Div = True
    m_typDetailCopy.Field01 = Trim(cboIcode.Text)
    m_typDetailCopy.Field02 = txtIname.Text
    m_typDetailCopy.Field03 = Format(imnQty.Value, "#,##0")
    m_typDetailCopy.Field04 = chkIdiv.Value
    m_typDetailCopy.Field05 = Format(imnPrice1.Value, "#,##0")
    m_typDetailCopy.Field06 = Format(imnPrice2.Value, "#,##0")
    m_typDetailCopy.Field07 = Format(imnPrice.Value, "#,##0")
    m_typDetailCopy.Field08 = Trim(cboBcode.Text)
    m_typDetailCopy.Field09 = lblBname.Caption
    m_typDetailCopy.Field10 = chkSdiv.Value
    m_typDetailCopy.Field11 = chkBdiv.Value
    m_typDetailCopy.Field12 = imnBnum.Value
    m_typDetailCopy.Field13 = imnSnum.Value
    
    Exit Sub

cmdCopy_Click_Err:

    Call MsgBox("明細コピークリック時エラー！！" _
            & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdCopy_Click_Err")

End Sub

'目　的　　：
'条　件　　：レコード移動クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub cmdDataMove_Click(Index As Integer)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo cmdDataMove_Click_Err

    Screen.MousePointer = vbHourglass

    With adoRecordset1
        strSQL = "SELECT * FROM DT030" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            Select Case Index
                Case 0:
                    .MoveFirst
                Case 1:
                    If Trim(txtOnum.Text) = "" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    .Find "Onum = " & Trim(txtOnum.Text)
                    If Not .EOF Then
                        .MovePrevious
                        If .EOF Or .BOF Then .MoveFirst
                    Else
                        .MoveFirst
                    End If
                Case 2:
                    If Trim(txtOnum.Text) = "" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    .Find "Onum = " & Trim(txtOnum.Text)
                    If Not .EOF Then
                        .MoveNext
                        If .EOF Or .BOF Then .MoveLast
                    Else
                        .MoveLast
                    End If
                Case 3:
                    .MoveLast
                Case Else
                    Exit Sub
            End Select
            Call FieldsClear(0)
            If FieldsSet(True, adoRecordset1) = False Then Exit Sub
        End If
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
cmdDataMove_Click_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("レコード移動クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdDataMove_Click_Err")
    
End Sub

Private Sub cmdDel_Click()

    If ListViewDelItem() = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    cboIcode_Kana.SetFocus

End Sub

Private Sub cmdDetail_Click()

    fraDetail.Visible = Not fraDetail.Visible

End Sub

Private Sub cmdEdit_Click()

    If DoValidationChecks_Dst() = False Then Exit Sub
    If ListViewSetItem(imnNo.Value, 0) = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    
    cboIcode_Kana.SetFocus

End Sub

'2005/08/12 追加
Private Sub cmdEtc_Click()
    
    If lblChargeDiv.Visible = True Then
        lblChargeDiv.Visible = False
        chkChargeDiv.Visible = False
        lblTaxDiv.Visible = False
        chkTaxDiv.Visible = False
        lblKeepDiv.Visible = False
        chkKeepDiv.Visible = False
        lblFixDiv.Visible = False
        chkFixDiv.Visible = False
    Else
        lblChargeDiv.Visible = True
        chkChargeDiv.Visible = True
        lblTaxDiv.Visible = True
        chkTaxDiv.Visible = True
        lblKeepDiv.Visible = True
        chkKeepDiv.Visible = True
        lblFixDiv.Visible = True
        chkFixDiv.Visible = True
    End If

End Sub

'目　的　　：
'条　件　　：実行クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub cmdExecute_Click()

    On Error Resume Next

    If MsgBox("実行しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub

    If optSyori(0).Value = True Or optSyori(1).Value = True Then
        '入力チェック
        If DoValidationChecks() = False Then Exit Sub
        If DataUpdate() = False Then Exit Sub
        m_strLastOnum = txtOnum.Text
        If optSyori(0).Value = True Then
            '2005/08/12 削除
'            If MsgBox("確認表を印刷しますか？", vbYesNo + vbQuestion, "") = vbYes Then
'               frmPrintDialog.m_blnAutoPrint = True
'                frmPrintDialog.Show vbModal
'            End If
        End If
    ElseIf optSyori(2).Value = True Then
        '入力チェック
        If DoValidationChecks() = False Then Exit Sub
        If DataDelete() = False Then Exit Sub
    End If
    
    'フィールドクリア
    Call FieldsClear(0)
    
    '2005/09/01 変更時は新規モードにする
    If optSyori(1).Value = True Then
        optSyori(0).Value = True
    ElseIf optSyori(0).Value = True And chkAutoCode.Value = 1 Then
        txtOnum.Text = AutoCodeSet
        txtOnum.SetFocus
    Else
        txtOnum.SetFocus
    End If
    
End Sub

'目　的　　：
'条　件　　：終了クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub cmdExit_Click()

    Unload Me
    End

End Sub

Private Sub cmdInsert_Click()

    Call SetImeMode(ActiveControl.hwnd, 2)
    If DoValidationChecks_Dst() = False Then Exit Sub
    If ListViewInsItem() = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    cboIcode_Kana.SetFocus

End Sub

Private Sub cmdLogin_Click()

    On Error GoTo cmdLogin_Click_Err

    g_blnLoginOK = False
    g_strPcode = Trim(cboPcode.Text)
    g_strPname = lblPname.Caption
    g_strOdate = Trim(lblOdate.Caption)
    frmLogin.Show vbModal
    If g_blnLoginOK = True Then
        lblOdate.Caption = g_strOdate
        cboPcode.Text = g_strPcode
        lblPname.Caption = g_strPname
        If optSyori(0).Value = True Then
            Call optSyori_Click(0)
        Else
            optSyori(0).Value = True
        End If
    End If
    Unload frmLogin
    m_strLastOnum = ""

    Exit Sub

cmdLogin_Click_Err:

    Call MsgBox("開催年月日と担当者の変更クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdLogin_Click_Err")

End Sub

'目　的　　：
'条　件　　：植木のマスタ登録クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub cmdMt050_Click()

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim lngCode As Long
    Dim blnAddNew As Boolean

    On Error GoTo cmdMt050_Click_Err

    If MsgBox("植木をマスタ登録しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    If Trim(txtIname.Text) = "" Then
        DoEvents
        Call MsgBox("植木名を入力してください。", vbOKOnly + vbCritical, "")
        txtIname.SetFocus
        DoEvents
        Exit Sub
    End If
    If Trim(cboIcode_Kana.Text) = "" Then
        DoEvents
        Call MsgBox("カナを入力してください。", vbOKOnly + vbCritical, "")
        cboIcode_Kana.SetFocus
        DoEvents
        Exit Sub
    End If
    
    blnAddNew = True
    With adoRecordset1
        '商品マスタ
        strSQL = "SELECT * FROM MT050" & _
                 " WHERE Iname = '" & Trim(txtIname.Text) & "'" & _
                 " ORDER BY Icode"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            '更新
            blnAddNew = False
            lngCode = .Fields("Icode")
            .Fields("Iname") = txtIname.Text
            .Fields("Ikana") = Global_LeftB_Ansi(cboIcode_Kana.Text, 40)
            .Update
        Else
            .Close
            blnAddNew = True
            '商品マスタ
            strSQL = "SELECT * FROM MT050" & _
                     " ORDER BY Icode DESC"
            .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If Not .EOF Then
                lngCode = CLng(.Fields("Icode")) + 1
            Else
                lngCode = 1
            End If
            .AddNew
            .Fields("Icode") = lngCode
            .Fields("Iname") = txtIname.Text
            .Fields("Ikana") = Global_LeftB_Ansi(cboIcode_Kana.Text, 40)
            .Fields("Idiv") = 1
            .Update
            .Close
        End If
    End With
        
    cboIcode.Text = lngCode
        
    txtIname.SetFocus
    DoEvents
    If blnAddNew Then
        Call MsgBox("新規登録しました。", vbOKOnly + vbInformation, "")
    Else
        Call MsgBox("更新登録しました。", vbOKOnly + vbInformation, "")
    End If
        
    Exit Sub

cmdMt050_Click_Err:

    Call MsgBox("植木のマスタ登録クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdMt050_Click_Err")

End Sub

'目　的　　：
'条　件　　：出品者のマスタ登録クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub cmdMt070_Click()

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim lngCode As Long
    Dim blnAddNew As Boolean

    On Error GoTo cmdMt070_Click_Err

    If MsgBox("出品者をマスタ登録しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
    
    If Trim(txtSname.Text) = "" Then
        DoEvents
        Call MsgBox("出品者名を入力してください。", vbOKOnly + vbCritical, "")
        txtSname.SetFocus
        DoEvents
        Exit Sub
    End If
    
    blnAddNew = True
    With adoRecordset1
        If Trim(cboScode.Text) = "" Or IsNumeric(cboScode.Text) = False Then
            '得意先マスタ
            strSQL = "SELECT * FROM MT070" & _
                     " WHERE Bname = '" & txtSname.Text & "'" & _
                     " ORDER BY Bcode"
        Else
            '得意先マスタ
            strSQL = "SELECT * FROM MT070" & _
                     " WHERE Bcode = " & cboScode.Text & _
                     " ORDER BY Bcode"
        End If
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            '更新
            blnAddNew = False
            lngCode = .Fields("Bcode")
            .Fields("Bcode") = lngCode
            .Fields("Bname") = txtSname.Text
            .Fields("Bkana") = Global_LeftB_Ansi(cboScode_Kana.Text, 20)
            .Fields("Addres") = txtAddres.Text
            .Fields("Tel") = txtTel.Text
            .Update
        Else
            .Close
            blnAddNew = True
            '得意先マスタ
            strSQL = "SELECT * FROM MT070" & _
                     " ORDER BY Bcode DESC"
            .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
            If Not .EOF Then
                lngCode = CLng(.Fields("Bcode")) + 1
            Else
                lngCode = 1
            End If
            .AddNew
            .Fields("Bcode") = lngCode
            .Fields("Bname") = txtSname.Text
            .Fields("Bkana") = Global_LeftB_Ansi(cboScode_Kana.Text, 20)
            .Fields("Addres") = txtAddres.Text
            .Fields("Tel") = txtTel.Text
            .Fields("Rname") = txtSname.Text
            .Fields("Rdiv") = RECEIPT_ON
            .Fields("Fdiv") = BUSINESS_DIV_EXHIBITION
            .Update
            .Close
        End If
    End With
        
    cboScode.Text = lngCode
        
    txtSname.SetFocus
    DoEvents
    If blnAddNew Then
        Call MsgBox("新規登録しました。", vbOKOnly + vbInformation, "")
    Else
        Call MsgBox("更新登録しました。", vbOKOnly + vbInformation, "")
    End If
        
    Exit Sub

cmdMt070_Click_Err:

    Call MsgBox("出品者のマスタ登録クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdMt070_Click_Err")

End Sub

Private Sub cmdPast_Click()

    On Error GoTo cmdPast_Click_Err

    If m_typDetailCopy.Div = False Then Exit Sub

    Call ListViewGetMaxRow
    cboIcode.Text = m_typDetailCopy.Field01
    txtIname.Text = m_typDetailCopy.Field02
    imnQty.Value = m_typDetailCopy.Field03
    chkIdiv.Value = m_typDetailCopy.Field04
    imnPrice1.Value = m_typDetailCopy.Field05
    imnPrice2.Value = m_typDetailCopy.Field06
    imnPrice.Value = m_typDetailCopy.Field07
    cboBcode.Text = m_typDetailCopy.Field08
    lblBname.Caption = m_typDetailCopy.Field09
'    chkSdiv.Value = m_typDetailCopy.Field10
'    chkBdiv.Value = m_typDetailCopy.Field11
'    imnBnum.Value = m_typDetailCopy.Field12
'    imnSnum.Value = m_typDetailCopy.Field13
    chkSdiv.Value = 0
    chkBdiv.Value = 0
    imnBnum.Value = 0
    imnSnum.Value = 0
    
    If DoValidationChecks_Dst() = False Then Exit Sub
    If ListViewSetItem(imnNo.Value, 0) = False Then Exit Sub
    Call Calc_Total
    Call FieldsClear(3)
    
    Exit Sub

cmdPast_Click_Err:

    Call MsgBox("明細貼付クリック時エラー！！" _
            & vbCrLf & Error$, vbOKOnly + vbCritical, "cmdPast_Click_Err")

End Sub

'目　的　　：
'条　件　　：検索クリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub cmdSearch_Click()

    frmSearch.Show vbModal

End Sub

Private Sub cmdSearch2_Click()

    frmSearch2.Show vbModal

End Sub

'目　的　　：
'条　件　　：フォームキーダウン時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
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
            cmdClear.SetFocus
            DoEvents
            Call cmdClear_Click
        Case vbKeyF9
            cmdExit.SetFocus
            DoEvents
            Call cmdExit_Click
        Case vbKeyF10
'            If cmdMt070.Enabled = False Then Exit Sub
'            cmdMt070.SetFocus
'            DoEvents
'            Call cmdMt070_Click
        Case vbKeyF11
            If cmdMt050.Enabled = False Then Exit Sub
            cmdMt050.SetFocus
            DoEvents
            Call cmdMt050_Click
        Case vbKeyF12
            cmdExecute.SetFocus
            DoEvents
            Call cmdExecute_Click
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

'目　的　　：
'条　件　　：フォームロード時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Err

    'フォームのセンタリング
    Me.left = (Screen.Width - Me.Width) / 2
    Me.top = ((Screen.Height - 450) - Me.Height) / 2

    Me.Caption = SYSTEM_NAME & "-" & "注文入力"

    lblOdate.Caption = g_strOdate
    cboPcode.Text = g_strPcode
    lblPname.Caption = g_strPname

    '処理ボタン
    optSyori(0).Value = True
    optSyori(1).Value = False
    optSyori(2).Value = False
    optSyori(3).Value = False
    optSyori(4).Value = False

    chkAutoCode.Value = AUTO_CODE
    If chkAutoCode.Value = 1 Then txtOnum.Text = AutoCodeSet()
    m_strLastOnum = ""
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("フォームロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")
    End

End Sub

'目　的　　：
'条　件　　：フォームアンロード時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Err
    
    Set g_clsAdoSQL = Nothing
    Set g_clsAdoAccess = Nothing
    Set g_clsReg = Nothing
    End
    
    Exit Sub
    
Form_Unload_Err:

    Call MsgBox("フォームアンロード時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")
    End

End Sub

'目　的　　：画面クリア
'条　件　　：
'結　果　　：
'引　数　　：0：全画面 1:キー部 2:明細部
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub FieldsClear(intKubun As Integer)

    On Error GoTo FieldsClear_Err
    
    If intKubun = 0 Then
        'キー
        txtOnum.Text = ""
        txtOnum.Tag = ""
        
        'ヘッダー
        txtPnum.Text = ""
        txtPnum.Tag = ""
                
        '2005/08/12 修正
'        chkChargeDiv.Value = 1
'        chkTaxDiv.Value = 1
'        chkKeepDiv.Value = 1
'        chkFixDiv.Value = 1
        
        chkChargeDiv.Value = 0
        chkTaxDiv.Value = 1
        chkKeepDiv.Value = 0
        chkFixDiv.Value = 1
        
        cboScode_Kana.Text = ""
        cboScode_Kana.Tag = ""
        txtSname.Text = ""
        cboScode.Text = ""
        cboScode.Tag = ""
        txtAddres.Text = ""
        txtTel.Text = ""
        optDiv(0).Value = True
        '明細
        imnNo.Value = 1
        cboIcode.Text = ""
        cboIcode.Tag = ""
        cboIcode_Kana.Text = ""
        txtIname.Text = ""
        imnQty.Value = 1            '数量の初期値「1」
        chkIdiv.Value = 0
        imnPrice1.Value = 0
        imnPrice2.Value = 0
        imnPrice.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
        'フッター
        lsvMeisai.ListItems.Clear
        imnQty_Total.Value = 0
        imnPrice_Total.Value = 0
        chkChumon.Value = 0
        fraChumon.Visible = False
        
        m_typDetailCopy.Div = False
    ElseIf intKubun = 1 Then
        'キー
        txtOnum.Text = ""
        txtOnum.Tag = ""
    ElseIf intKubun = 2 Then
        'ヘッダー
        txtPnum.Text = ""
        txtPnum.Tag = ""
        chkChargeDiv.Value = 1
        chkTaxDiv.Value = 1
        chkKeepDiv.Value = 1
        chkFixDiv.Value = 1
        cboScode_Kana.Text = ""
        cboScode_Kana.Tag = ""
        txtSname.Text = ""
        cboScode.Text = ""
        cboScode.Tag = ""
        txtAddres.Text = ""
        txtTel.Text = ""
        optDiv(0).Value = True
        '明細
        imnNo.Value = 1
        cboIcode.Text = ""
        cboIcode.Tag = ""
        cboIcode_Kana.Text = ""
        txtIname.Text = ""
        imnQty.Value = 1            '数量の初期値「1」
        imnPrice1.Value = 0
        imnPrice2.Value = 0
        imnPrice.Value = 0
        chkIdiv.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
        'フッター
        lsvMeisai.ListItems.Clear
        imnQty_Total.Value = 0
        imnPrice_Total.Value = 0
        chkChumon.Value = 0
        fraChumon.Visible = False
        
        m_typDetailCopy.Div = False
    ElseIf intKubun = 3 Then
        '明細
        cboIcode.Text = ""
        cboIcode_Kana.Text = ""
        txtIname.Text = ""
        imnQty.Value = 1            '数量の初期値「1」
        chkIdiv.Value = 0
        imnPrice1.Value = 0
        imnPrice2.Value = 0
        imnPrice.Value = 0
        cboBcode.Text = ""
        cboBcode.Tag = ""
        lblBname.Caption = ""
        chkSdiv.Value = 0
        chkBdiv.Value = 0
        imnBnum.Value = 0
        imnSnum.Value = 0
    End If
        
    chkChumon.Value = 1
        
    Exit Sub
    
FieldsClear_Err:

    Call MsgBox("画面クリアエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsClear_Err")

End Sub

Private Sub imnBnum_GotFocus()
    
    imnBnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnBnum_LostFocus()
    
    imnBnum.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnPrice1_GotFocus()
    
    imnPrice1.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnPrice1_LostFocus()
    
    imnPrice1.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnPrice1_Validate(Cancel As Boolean)

    If imnPrice2.Value > imnPrice1.Value Then
        If MsgBox("仕入単価のほうが高いですが、よろしいですか？", vbYesNo + vbInformation, "") = vbNo Then
            Cancel = True
        End If
    End If

    If Calc_Price() = False Then Cancel = True
    
End Sub

Private Sub imnPrice2_GotFocus()
    
    imnPrice2.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnPrice2_LostFocus()
    
    imnPrice2.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imnPrice2_Validate(Cancel As Boolean)

'    If imnPrice2.Value > imnPrice1.Value Then
'        If MsgBox("仕入単価のほうが高いですが、よろしいですか？", vbYesNo + vbInformation, "") = vbNo Then
'            Cancel = True
'        End If
'    End If

    If Calc_Price() = False Then Cancel = True
    
End Sub

Private Sub imnQty_GotFocus()
    
    imnQty.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub imnQty_LostFocus()
    
    imnQty.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub imnQty_Validate(Cancel As Boolean)

    If Calc_Price() = False Then Cancel = True

End Sub

Private Sub imnSnum_GotFocus()
    
    imnSnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub imnSnum_LostFocus()
    
    imnSnum.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub imtIcode_Kana_Focus1_GotFocus()

'    On Error Resume Next

'    If Trim(txtIname.Text) = "" Then
        txtIname.SetFocus
'    Else
'        imnQty.SetFocus
'    End If

End Sub

Private Sub imtIcode_Kana_Focus2_GotFocus()

    cboIcode_Kana.SetFocus

End Sub

Private Sub imtScode_Kana_Focus1_GotFocus()

    If Trim(txtSname.Text) = "" Then
        txtSname.SetFocus
    Else
        If optSyori(0).Value = True Or optSyori(1).Value = True Then
            cboIcode_Kana.SetFocus
        Else
            cmdExecute.SetFocus
        End If
    End If

End Sub

Private Sub imtScode_Kana_Focus2_GotFocus()

    txtTel.SetFocus

End Sub

Private Sub imtFocusEnd_GotFocus()

    cmdExit.SetFocus

End Sub

Private Sub imtFocusFirst_GotFocus()

    If optSyori(0).Value = True Then
        If txtSname.Enabled = True Then
            txtSname.SetFocus
        Else
            txtOnum.SetFocus
        End If
    Else
        txtOnum.SetFocus
    End If

End Sub

Private Sub lsvMeisai_Click()

    On Error Resume Next

    '行が選択されているか？
    If lsvMeisai.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    '明細表示
    Call ListViewGetItem
    
    cboIcode_Kana.SetFocus

End Sub

'目　的　　：
'条　件　　：処理区分ボタンクリック時
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub optSyori_Click(Index As Integer)

    Dim intIndex1 As Integer
    Dim adoRecordset1 As New ADODB.Recordset

    On Error GoTo optSyori_Click_Err

    '画面クリア
    Call FieldsClear(0)
    
    '背景色の変更
    For intIndex1 = 0 To 5
        If intIndex1 = Index Then
            optSyori(intIndex1).BackColor = BUTTON_ON
        Else
            optSyori(intIndex1).BackColor = BUTTON_OFF
        End If
    Next intIndex1
    
    Select Case Index
        Case 0: '新規
            Call FieldsControl(0, True)
            Call FieldsControl(1, False)
            Call FieldsControl(2, True)
            Call FieldsControl(3, True)
            If chkAutoCode.Value = 1 Then txtOnum.Text = AutoCodeSet
        Case 1: '変更
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
            Call FieldsControl(3, True)
        Case 2: '削除
            Call FieldsControl(0, True)
            Call FieldsControl(1, True)
            Call FieldsControl(2, False)
            Call FieldsControl(3, False)
        Case 3: '印刷
            Call FieldsControl(0, True)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
            Call FieldsControl(3, False)
            frmPrintDialog.m_blnAutoPrint = False
            frmPrintDialog.Show vbModal
        Case 4: '伝票印刷
            Call FieldsControl(0, True)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
            Call FieldsControl(3, False)
            frmPrintDialog2.m_blnAutoPrint = False
            frmPrintDialog2.Show vbModal
        Case 5: '集計表印刷
            Call FieldsControl(0, True)
            Call FieldsControl(1, False)
            Call FieldsControl(2, False)
            Call FieldsControl(3, False)
            frmPrintDialog3.Show vbModal
    End Select

    On Error Resume Next
    txtOnum.SetFocus
    DoEvents
    
    Exit Sub

optSyori_Click_Err:

    Call MsgBox("処理区分クリック時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Unload_Err")

End Sub

Private Sub txtAddres_GotFocus()

    txtAddres.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtAddres_LostFocus()

    txtAddres.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtIname_GotFocus()

    txtIname.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtIname_LostFocus()

    txtIname.BackColor = FOCUS_NO_COLOR
    
End Sub

Private Sub txtOnum_Change()

    If Trim(txtOnum.Text) = "" Then Exit Sub

    If txtOnum.Tag <> txtOnum.Text Then
        If optSyori(0).Value Or optSyori(1).Value Then
            fraMeisai.Enabled = True
            DoEvents
        End If
    End If
    
End Sub

Private Sub txtOnum_GotFocus()

    txtOnum.Tag = txtOnum.Text
    txtOnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtOnum_LostFocus()

    txtOnum.Tag = ""
    txtOnum.BackColor = FOCUS_NO_COLOR

End Sub

Private Sub txtOnum_Validate(Cancel As Boolean)

    If Trim(txtOnum.Text) = "" Then Exit Sub
    If txtOnum.Tag = txtOnum.Text Then Exit Sub

    If optSyori(0).Value = True Then
        If FieldsSet(False) = True Then
            Cancel = True
            Call MsgBox("既にデータが存在します。", vbOKOnly + vbCritical, "")
            Exit Sub
        End If
    Else
        If FieldsSet(True) = False Then
            Cancel = True
            Call MsgBox("データが存在しません。", vbOKOnly + vbCritical, "")
            Exit Sub
        End If
    End If

End Sub

Private Sub txtPnum_GotFocus()
    
    txtPnum.BackColor = FOCUS_STOP_COLOR
    
End Sub

Private Sub txtPnum_LostFocus()

    txtPnum.BackColor = FOCUS_NO_COLOR

End Sub

'目　的　　：入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function DoValidationChecks() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Err

    If Trim(lblOdate.Caption) = "" Then
        strErrMsg = "開催年月日を入力してください。"
        GoTo ErrorTrap:
    End If
    If Trim(txtOnum.Text) = "" Then
        strErrMsg = "注文番号を入力してください。"
        txtOnum.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(txtSname.Text) = "" Then
        strErrMsg = "出品者名を入力してください。"
        txtSname.SetFocus
        GoTo ErrorTrap:
    End If
    If Global_LenB_Ansi(Trim(txtSname.Text)) > 40 Then
        strErrMsg = "出品者名の文字数が多すぎます。"
        txtSname.SetFocus
        GoTo ErrorTrap:
    End If
    If lsvMeisai.ListItems.Count <= 0 Then
        strErrMsg = "明細を入力してください。"
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

'目　的　　：フィールドの制御
'条　件　　：
'結　果　　：
'引　数　　：intKbn 0:キー部 1:レコード移動 2:明細　3:ヘッダー
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub FieldsControl(intKbn As Integer, blnEnabled As Boolean)

    On Error GoTo FieldsControl_Err
    
    Select Case intKbn
        Case 0:
            fraKey.Enabled = blnEnabled
        Case 1:
            fraRecordSelector.Enabled = blnEnabled
            cmdDataMove(0).Enabled = blnEnabled
            cmdDataMove(1).Enabled = blnEnabled
            cmdDataMove(2).Enabled = blnEnabled
            cmdDataMove(3).Enabled = blnEnabled
            cmdSearch.Enabled = blnEnabled
        Case 2:
            fraMeisai.Enabled = blnEnabled
        Case 3:
            fraHeader.Enabled = blnEnabled
    End Select
        
    Exit Sub

FieldsControl_Err:

    Call MsgBox("フィールドの制御エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsControl_Err")

End Sub

'目　的　　：フィールドのセット
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Public Function FieldsSet(blnVisible As Boolean, Optional adoRecordsetArg As Variant) As Boolean
    
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim strSQL As String
    Dim itmX As ListItem
    Dim intIndex1 As Integer
    Dim intindex2 As Integer
    Dim strBuff As String
    Dim varColor As Variant

    On Error GoTo FieldsSet_Err
    
    FieldsSet = False
    
    Screen.MousePointer = vbHourglass
    
    If IsMissing(adoRecordsetArg) = False Then
        Set adoRecordset1 = adoRecordsetArg
    Else
        strSQL = "SELECT * FROM DT030" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Onum = " & txtOnum.Text
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    End If
    
    With adoRecordset1
        If .EOF Or .BOF Then
            adoRecordset1.Close
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        If blnVisible = False Then
            adoRecordset1.Close
            FieldsSet = True
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        txtOnum.Text = IIf(IsNull(.Fields("Onum")), "", .Fields("Onum"))
        txtPnum.Text = IIf(IsNull(.Fields("Pnum")), "", .Fields("Pnum"))
        cboScode.Text = IIf(IsNull(.Fields("Scode")), "", Trim(.Fields("Scode")))
        txtSname.Text = IIf(IsNull(.Fields("Sname")), "", Trim(.Fields("Sname")))
        txtAddres.Text = IIf(IsNull(.Fields("Addres")), "", Trim(.Fields("Addres")))
        txtTel.Text = IIf(IsNull(.Fields("Tel")), "", Trim(.Fields("Tel")))
        If Not IsNull(.Fields("Div")) Then
            If .Fields("Div") = TIKU_DIV_OFF Then
                optDiv(1).Value = True
            ElseIf .Fields("Div") = TIKU_DIV_ON Then
                optDiv(0).Value = True
            End If
        End If
        chkSoukin.Value = IIf(IsNull(.Fields("Soukin")), 0, .Fields("Soukin"))
        chkChargeDiv.Value = IIf(IsNull(.Fields("ChargeDiv")), 0, .Fields("ChargeDiv"))
        chkTaxDiv.Value = IIf(IsNull(.Fields("TaxDiv")), 0, .Fields("TaxDiv"))
        chkKeepDiv.Value = IIf(IsNull(.Fields("KeepDiv")), 0, .Fields("KeepDiv"))
        chkFixDiv.Value = IIf(IsNull(.Fields("FixDiv")), 0, .Fields("FixDiv"))
        .Close
    End With
    
    With adoRecordset2
        intIndex1 = 1
        lsvMeisai.ListItems.Clear
        
        strSQL = "SELECT * FROM DT031" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Onum = " & txtOnum.Text & _
                 " ORDER BY Odate,Onum,Line"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Set itmX = lsvMeisai.ListItems.Add(, , intIndex1, 0)
            itmX.SubItems(1) = IIf(IsNull(.Fields("Icode")), "", .Fields("Icode"))
            itmX.SubItems(2) = IIf(IsNull(.Fields("Iname")), "", Trim(.Fields("Iname")))
            itmX.SubItems(3) = IIf(IsNull(.Fields("Qty")), 0, Format(.Fields("Qty"), "#,##0"))
            itmX.SubItems(4) = IIf(IsNull(.Fields("Idiv")), 0, .Fields("Idiv"))

'            itmX.SubItems(5) = IIf(IsNull(.Fields("Price1")), "", Format(.Fields("Price1"), "#,##0"))
'            itmX.SubItems(6) = IIf(IsNull(.Fields("Price2")), "", Format(.Fields("Price2"), "#,##0"))
            itmX.SubItems(6) = IIf(IsNull(.Fields("Price1")), "", Format(.Fields("Price1"), "#,##0"))
            itmX.SubItems(5) = IIf(IsNull(.Fields("Price2")), "", Format(.Fields("Price2"), "#,##0"))
            
            itmX.SubItems(7) = IIf(IsNull(.Fields("Price")), "", Format(.Fields("Price"), "#,##0"))
            itmX.SubItems(8) = IIf(IsNull(.Fields("Bcode")), "", .Fields("Bcode"))
            If Not IsNull(.Fields("Bcode")) Then
                itmX.SubItems(9) = IIf(IsNull(.Fields("Bcode")), "", Global_Get_Bname(g_clsAdoSQL, .Fields("Bcode"), lblOdate.Caption, strBuff))
            Else
                itmX.SubItems(9) = ""
            End If
            itmX.SubItems(10) = IIf(IsNull(.Fields("Sdiv")), 0, .Fields("Sdiv"))
            itmX.SubItems(11) = IIf(IsNull(.Fields("Bdiv")), 0, .Fields("Bdiv"))
            itmX.SubItems(12) = IIf(IsNull(.Fields("Bnum")), 0, .Fields("Bnum"))
            itmX.SubItems(13) = IIf(IsNull(.Fields("Snum")), 0, .Fields("Snum"))
            
'            '入力済みの場合は、前景色を変える
'            If itmX.SubItems(4) = INPUT_ON Or itmX.SubItems(7) <> "0" Then
'                varColor = DETAIL_FORECOLOR2
'            Else
'                varColor = DETAIL_FORECOLOR1
'            End If
'            itmX.ForeColor = varColor
'            For intIndex2 = 1 To MAX_COL
'                itmX.ListSubItems(intIndex2).ForeColor = varColor
'            Next intIndex2

            intIndex1 = intIndex1 + 1
            .MoveNext
        Loop
        .Close
        
        Call ListViewGetMaxRow
    End With

    Call Calc_Total
    
    Set adoRecordset1 = Nothing
    Set adoRecordset2 = Nothing
    
    Screen.MousePointer = vbDefault
    
    FieldsSet = True
    
    Exit Function

FieldsSet_Err:

    FieldsSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("フィールドセットエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "FieldsSet_Err")

End Function

'目　的　　：データの登録
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function DataUpdate() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim intIndex1 As Integer
    Dim intindex2 As Integer
    Dim lngCount As Long
    Dim lngDeleteCount As Long

    On Error GoTo DataUpdate_Err
    
    Screen.MousePointer = vbHourglass
    
    lngCount = 0
    lngDeleteCount = 0
    
    g_clsAdoSQL.Connection.BeginTrans
    
    If optSyori(0).Value = True Then
        If chkAutoCode.Value = 1 Then txtOnum.Text = AutoCodeSet()
    ElseIf optSyori(1).Value = True Then
        
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            '注文明細データ
            strSQL = "SELECT * FROM DT031" & _
                     " WHERE Odate = '" & Trim(lblOdate.Caption) & "'" & _
                     " And Onum = " & txtOnum.Text & _
                     " And Line = " & intIndex1
            adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
            If Not adoRecordset1.EOF Then
                If Global_StrNull(adoRecordset1.Fields("Bdiv")) = BUYER_REPORT_ON Or Global_StrNull(adoRecordset1.Fields("Bnum")) <> 0 Then
                    '買主が違う場合
                    If Global_StrNull(lsvMeisai.ListItems(intIndex1).SubItems(8)) <> Global_StrNull(adoRecordset1.Fields("Bcode")) Then
                        If Not IsNull(adoRecordset1.Fields("Bcode")) Then
                
                        '********** 変更前データ **********
                                                        
                            '買主精算データを削除
                            strSQL = "DELETE FROM DT041" & _
                                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                                     " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                                     " AND Num >= " & adoRecordset1.Fields("Bnum")
                            g_clsAdoSQL.Connection.Execute strSQL, lngCount
                    
                            '競売明細データを更新
                            strSQL = "UPDATE DT021" & _
                                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                                     " Bnum = 0 " & _
                                     " WHERE LEFT(Ocode, 8) = '" & Format$(lblOdate.Caption, "yyyymmdd") & "'" & _
                                     " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                                     " AND BNum >= " & adoRecordset1.Fields("Bnum")
                            g_clsAdoSQL.Connection.Execute strSQL, lngCount
                    
                            '注文データを更新
                            strSQL = "UPDATE DT031" & _
                                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                                     " Bnum = 0 " & _
                                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                                     " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                                     " AND BNum >= " & adoRecordset1.Fields("Bnum")
                            g_clsAdoSQL.Connection.Execute strSQL, lngCount
                        
                            '注文データ画面ワークを更新
                            For intindex2 = 1 To lsvMeisai.ListItems.Count
'                                If Global_StrNull(lsvMeisai.ListItems(intindex2).SubItems(8)) = adoRecordset1.Fields("Bcode") And _
                                   lsvMeisai.ListItems(intindex2).SubItems(12) = adoRecordset1.Fields("Bnum") Then
                                '買主コードの比較
                                If Global_StrNull(lsvMeisai.ListItems(intindex2).SubItems(8)) = adoRecordset1.Fields("Bcode") Then
                                    '買主精算回数とフラグのクリア
                                    lsvMeisai.ListItems(intindex2).SubItems(11) = 0
                                    lsvMeisai.ListItems(intindex2).SubItems(12) = 0
                                End If
                            Next intindex2
                        
                        '********** 変更後データ **********
                                                        
                            '買主精算データを削除
                            strSQL = "DELETE FROM DT041" & _
                                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                                     " AND Bcode = " & lsvMeisai.ListItems(intIndex1).SubItems(8) & _
                                     " AND Num >= " & lsvMeisai.ListItems(intIndex1).SubItems(12)
                            g_clsAdoSQL.Connection.Execute strSQL, lngCount
                    
                            '競売明細データを更新
                            strSQL = "UPDATE DT021" & _
                                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                                     " Bnum = 0 " & _
                                     " WHERE LEFT(Ocode, 8) = '" & Format$(lblOdate.Caption, "yyyymmdd") & "'" & _
                                     " AND Bcode = " & lsvMeisai.ListItems(intIndex1).SubItems(8) & _
                                     " AND BNum >= " & lsvMeisai.ListItems(intIndex1).SubItems(12)
                            g_clsAdoSQL.Connection.Execute strSQL, lngCount
                    
                            '注文データを更新
                            strSQL = "UPDATE DT031" & _
                                     " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                                     " Bnum = 0 " & _
                                     " WHERE Odate = '" & lblOdate.Caption & "'" & _
                                     " AND Bcode = " & lsvMeisai.ListItems(intIndex1).SubItems(8) & _
                                     " AND BNum >= " & lsvMeisai.ListItems(intIndex1).SubItems(12)
                            g_clsAdoSQL.Connection.Execute strSQL, lngCount
                        
                            '注文データ画面ワークを更新
                            For intindex2 = 1 To lsvMeisai.ListItems.Count
                                '現在の行は更新しない
                                If intIndex1 <> intindex2 Then
'                                    If Global_StrNull(lsvMeisai.ListItems(intindex2).SubItems(8)) = lsvMeisai.ListItems(intIndex1).SubItems(8) And _
                                       lsvMeisai.ListItems(intindex2).SubItems(12) = lsvMeisai.ListItems(intIndex1).SubItems(12) Then
                                    '買主コードの比較
                                    If Global_StrNull(lsvMeisai.ListItems(intindex2).SubItems(8)) = lsvMeisai.ListItems(intIndex1).SubItems(8) Then
                                        '買主精算回数とフラグのクリア
                                        lsvMeisai.ListItems(intindex2).SubItems(11) = 0
                                        lsvMeisai.ListItems(intindex2).SubItems(12) = 0
                                    End If
                                End If
                            Next intindex2
                        
                        '********** 現在行 **********
                            
                            '買主精算回数とフラグのクリア
                            lsvMeisai.ListItems(intIndex1).SubItems(11) = 0
                            lsvMeisai.ListItems(intIndex1).SubItems(12) = 0
                        
                            lngDeleteCount = lngDeleteCount + 1
                        
                        End If
                    End If
                    
                    '金額が違う場合
                    If CCur(lsvMeisai.ListItems(intIndex1).SubItems(7)) <> CCur(adoRecordset1.Fields("Price")) Then
                        '********** 変更前データ **********
                                                        
                        '買主精算データを削除
                        strSQL = "DELETE FROM DT041" & _
                                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                                 " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                                 " AND Num >= " & adoRecordset1.Fields("Bnum")
                        g_clsAdoSQL.Connection.Execute strSQL, lngCount
                
                        '競売明細データを更新
                        strSQL = "UPDATE DT021" & _
                                 " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                                 " Bnum = 0 " & _
                                 " WHERE LEFT(Ocode, 8) = '" & Format$(lblOdate.Caption, "yyyymmdd") & "'" & _
                                 " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                                 " AND BNum >= " & adoRecordset1.Fields("Bnum")
                        g_clsAdoSQL.Connection.Execute strSQL, lngCount
                
                        '注文データを更新
                        strSQL = "UPDATE DT031" & _
                                 " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                                 " Bnum = 0 " & _
                                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                                 " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                                 " AND BNum >= " & adoRecordset1.Fields("Bnum")
                        g_clsAdoSQL.Connection.Execute strSQL, lngCount
                    
                        '注文データ画面ワークを更新
                        For intindex2 = 1 To lsvMeisai.ListItems.Count
'                            If Global_StrNull(lsvMeisai.ListItems(intindex2).SubItems(8)) = adoRecordset1.Fields("Bcode") And _
                               lsvMeisai.ListItems(intindex2).SubItems(12) = adoRecordset1.Fields("Bnum") Then
                            '買主コードの比較
                            If Global_StrNull(lsvMeisai.ListItems(intindex2).SubItems(8)) = adoRecordset1.Fields("Bcode") Then
                                '買主精算回数とフラグのクリア
                                lsvMeisai.ListItems(intindex2).SubItems(11) = 0
                                lsvMeisai.ListItems(intindex2).SubItems(12) = 0
                            End If
                        Next intindex2
                    
                    '********** 変更後データ **********
                                                    
                        '買主精算データを削除
                        strSQL = "DELETE FROM DT041" & _
                                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                                 " AND Bcode = " & lsvMeisai.ListItems(intIndex1).SubItems(8) & _
                                 " AND Num >= " & lsvMeisai.ListItems(intIndex1).SubItems(12)
                        g_clsAdoSQL.Connection.Execute strSQL, lngCount
                
                        '競売明細データを更新
                        strSQL = "UPDATE DT021" & _
                                 " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                                 " Bnum = 0 " & _
                                 " WHERE LEFT(Ocode, 8) = '" & Format$(lblOdate.Caption, "yyyymmdd") & "'" & _
                                 " AND Bcode = " & lsvMeisai.ListItems(intIndex1).SubItems(8) & _
                                 " AND BNum >= " & lsvMeisai.ListItems(intIndex1).SubItems(12)
                        g_clsAdoSQL.Connection.Execute strSQL, lngCount
                
                        '注文データを更新
                        strSQL = "UPDATE DT031" & _
                                 " SET Bdiv = " & BUYER_REPORT_OFF & "," & _
                                 " Bnum = 0 " & _
                                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                                 " AND Bcode = " & lsvMeisai.ListItems(intIndex1).SubItems(8) & _
                                 " AND BNum >= " & lsvMeisai.ListItems(intIndex1).SubItems(12)
                        g_clsAdoSQL.Connection.Execute strSQL, lngCount
                    
                        '注文データ画面ワークを更新
                        For intindex2 = 1 To lsvMeisai.ListItems.Count
                            '現在の行は更新しない
                            If intIndex1 <> intindex2 Then
'                                If Global_StrNull(lsvMeisai.ListItems(intindex2).SubItems(8)) = lsvMeisai.ListItems(intIndex1).SubItems(8) And _
                                   lsvMeisai.ListItems(intindex2).SubItems(12) = lsvMeisai.ListItems(intIndex1).SubItems(12) Then
                                '買主コードの比較
                                If Global_StrNull(lsvMeisai.ListItems(intindex2).SubItems(8)) = lsvMeisai.ListItems(intIndex1).SubItems(8) Then
                                    '買主精算回数とフラグのクリア
                                    lsvMeisai.ListItems(intindex2).SubItems(11) = 0
                                    lsvMeisai.ListItems(intindex2).SubItems(12) = 0
                                End If
                            End If
                        Next intindex2
                    
                    '********** 現在行 **********
                        
                        '買主精算回数とフラグのクリア
                        lsvMeisai.ListItems(intIndex1).SubItems(11) = 0
                        lsvMeisai.ListItems(intIndex1).SubItems(12) = 0
                    
                        lngDeleteCount = lngDeleteCount + 1
                    
                    End If
                End If
            End If
            adoRecordset1.Close
        Next
    End If
 
    strSQL = "DELETE FROM DT031" & _
             " WHERE Odate = '" & lblOdate.Caption & "'" & _
             " AND Onum = " & txtOnum.Text
    g_clsAdoSQL.Connection.Execute strSQL
    
    strSQL = "DELETE FROM DT030" & _
             " WHERE Odate = '" & lblOdate.Caption & "'" & _
             " AND Onum = " & txtOnum.Text
    g_clsAdoSQL.Connection.Execute strSQL
 
    With adoRecordset1
        strSQL = "SELECT * FROM DT030" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Onum = " & txtOnum.Text
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        If .EOF Then .AddNew
        .Fields("Odate") = lblOdate.Caption
        .Fields("Onum") = txtOnum.Text
        .Fields("Pnum") = IIf(Trim(txtPnum.Text) = "", Null, txtPnum.Text)
        If IsNumeric(cboScode.Text) = True Then .Fields("Scode") = cboScode.Text
        .Fields("Sname") = Trim(txtSname.Text)
        .Fields("Addres") = txtAddres.Text
        .Fields("Tel") = txtTel.Text
        If optDiv(0).Value = True Then
            .Fields("Div") = TIKU_DIV_ON
        ElseIf optDiv(1).Value = True Then
            .Fields("Div") = TIKU_DIV_OFF
        End If
        .Fields("Soukin") = chkSoukin.Value
        .Fields("ChargeDiv") = chkChargeDiv.Value
        .Fields("TaxDiv") = chkTaxDiv.Value
        .Fields("KeepDiv") = chkKeepDiv.Value
        .Fields("FixDiv") = chkFixDiv.Value
        .Update
        .Close
    End With
    
    With adoRecordset2
        strSQL = "SELECT * FROM DT031"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockOptimistic
        For intIndex1 = 1 To lsvMeisai.ListItems.Count
            .AddNew
            .Fields("Odate") = lblOdate.Caption
            .Fields("Onum") = txtOnum.Text
            .Fields("Line") = intIndex1
            .Fields("Icode") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(1)), lsvMeisai.ListItems(intIndex1).SubItems(1), Null)
            .Fields("Iname") = lsvMeisai.ListItems(intIndex1).SubItems(2)
            .Fields("Qty") = lsvMeisai.ListItems(intIndex1).SubItems(3)
'            .Fields("Price1") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(5)), lsvMeisai.ListItems(intIndex1).SubItems(5), Null)
'            .Fields("Price2") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(6)), lsvMeisai.ListItems(intIndex1).SubItems(6), Null)
            .Fields("Price1") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(6)), lsvMeisai.ListItems(intIndex1).SubItems(6), Null)
            .Fields("Price2") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(5)), lsvMeisai.ListItems(intIndex1).SubItems(5), Null)
            
            .Fields("Price") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(7)), lsvMeisai.ListItems(intIndex1).SubItems(7), Null)
            .Fields("Bcode") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(8)), lsvMeisai.ListItems(intIndex1).SubItems(8), Null)
            .Fields("Sdiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(10)), lsvMeisai.ListItems(intIndex1).SubItems(10), 0)
            .Fields("Bdiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(11)), lsvMeisai.ListItems(intIndex1).SubItems(11), 0)
            .Fields("Bnum") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(12)), lsvMeisai.ListItems(intIndex1).SubItems(12), 0)
            .Fields("Snum") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(13)), lsvMeisai.ListItems(intIndex1).SubItems(13), 0)
            .Fields("Idiv") = IIf(IsNumeric(lsvMeisai.ListItems(intIndex1).SubItems(4)), lsvMeisai.ListItems(intIndex1).SubItems(4), 0)
            .Update
        Next
        .Close
    End With
    
    g_clsAdoSQL.Connection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    Set adoRecordset1 = Nothing
    Set adoRecordset2 = Nothing
    
    If lngDeleteCount > 0 Then
        Call MsgBox("データが変更されました。もう一度買主精算を行ってください。" _
                    , vbOKOnly + vbExclamation, "注意")
    End If
    
    DataUpdate = True
    
    Exit Function

DataUpdate_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    DataUpdate = False
    Screen.MousePointer = vbDefault
    Call MsgBox("データ登録エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataUpdate_Err")

End Function

'目　的　　：データの削除
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function DataDelete() As Boolean

    Dim strSQL As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim lngCount As Long
    Dim lngDeleteCount As Long
    
    On Error GoTo DataDelete_Err
    
    Screen.MousePointer = vbHourglass
    
'    '出品者精算データ
'    strSQL = "SELECT * FROM DT040" & _
'             " WHERE Odate = '" & Trim(lblOdate.Caption) & "'" & _
'             " AND Pnum = " & txtPnum.Text & _
'             " ORDER BY Odate,Pnum,Num DESC"
'    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
'    If adoRecordset1.EOF = False Then
'        DataDelete = False
'        Screen.MousePointer = vbDefault
'        Call MsgBox("既に精算されているため削除できません。", vbOKOnly + vbCritical, "")
'        Exit Function
'    End If
'    adoRecordset1.Close
    
    With g_clsAdoSQL.Connection
        .BeginTrans
        
        lngCount = 0
        lngDeleteCount = 0
        
        '注文明細データ
        strSQL = "SELECT * FROM DT031" & _
                 " WHERE Odate = '" & Trim(lblOdate.Caption) & "'" & _
                 " And Onum = " & txtOnum.Text & _
                 " ORDER BY Odate,Onum,Line"
        adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not adoRecordset1.EOF
            If Global_StrNull(adoRecordset1.Fields("Bdiv")) = BUYER_REPORT_ON Or Global_StrNull(adoRecordset1.Fields("Bnum")) <> 0 Then
        
                '買主精算データを削除
                strSQL = "DELETE FROM DT041" & _
                         " WHERE Odate = '" & lblOdate.Caption & "'" & _
                         " AND Bcode = " & adoRecordset1.Fields("Bcode") & _
                         " AND Num = " & adoRecordset1.Fields("Bnum")
                g_clsAdoSQL.Connection.Execute strSQL, lngCount
            
                lngDeleteCount = lngDeleteCount + lngCount
                    
            End If
            adoRecordset1.MoveNext
        Loop
        adoRecordset1.Close
        
        strSQL = "DELETE FROM DT031" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Onum = " & txtOnum.Text
        .Execute strSQL
    
        strSQL = "DELETE FROM DT030" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " AND Onum = " & txtOnum.Text
        .Execute strSQL
    
        .CommitTrans
    End With
    
    Screen.MousePointer = vbDefault
    
    If lngDeleteCount > 0 Then
        Call MsgBox("買主精算データが削除されました。もう一度買主精算を行ってください。" _
                    , vbOKOnly + vbExclamation, "注意")
    End If
    
    DataDelete = True
    
    Exit Function

DataDelete_Err:

    g_clsAdoSQL.Connection.RollbackTrans
    DataDelete = False
    Screen.MousePointer = vbDefault
    Call MsgBox("データの削除エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DataDelete_Err")

End Function

'目　的　　：コードの自動採番
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function AutoCodeSet() As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo AutoCodeSet_Err
    
    AutoCodeSet = ""
    
    Screen.MousePointer = vbHourglass
    
    With adoRecordset1
        strSQL = "SELECT Onum FROM DT030" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " ORDER BY Odate,Onum"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If .EOF Or .BOF Then
            AutoCodeSet = 1
            adoRecordset1.Close
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        .MoveLast
        If CLng(.Fields("Onum")) < 9999 Then
            AutoCodeSet = CLng(.Fields("Onum")) + 1
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Function

AutoCodeSet_Err:

    AutoCodeSet = False
    Screen.MousePointer = vbDefault
    Call MsgBox("コードの自動採番エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "AutoCodeSet_Err")

End Function

'目　的　　：リストビューへのデータ登録
'条　件　　：
'結　果　　：
'引　数　　：intFlg(0:追加・更新 1:挿入)
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function ListViewSetItem(intPostion As Integer, intFlg As Integer) As Boolean

    Dim itmX As ListItem
    Dim intIndex1 As Integer
    Dim varColor As Variant

    On Error GoTo ListViewSetItem_Err
    
    ListViewSetItem = False
    
    'リストビューのデータ検索（行番号が一致するデータがあったら削除）
    Set itmX = lsvMeisai.FindItem(intPostion, , , 0)
    If Not (itmX Is Nothing) Then
        If intFlg = 0 Then
            'データ削除
            lsvMeisai.ListItems.Remove itmX.Index
        End If
        'データを追加
        Set itmX = lsvMeisai.ListItems.Add(intPostion, , intPostion, 0)
    Else
        'データを追加
        Set itmX = lsvMeisai.ListItems.Add(, , intPostion, 0)
    End If
    itmX.SubItems(1) = Trim(cboIcode.Text)
    itmX.SubItems(2) = txtIname.Text
    itmX.SubItems(3) = Format(imnQty.Value, "#,##0")
    itmX.SubItems(4) = chkIdiv.Value

'    itmX.SubItems(5) = Format(imnPrice1.Value, "#,##0")
'    itmX.SubItems(6) = Format(imnPrice2.Value, "#,##0")
    itmX.SubItems(6) = Format(imnPrice1.Value, "#,##0")
    itmX.SubItems(5) = Format(imnPrice2.Value, "#,##0")
    
    itmX.SubItems(7) = Format(imnPrice.Value, "#,##0")
    itmX.SubItems(8) = Trim(cboBcode.Text)
    itmX.SubItems(9) = lblBname.Caption
    itmX.SubItems(10) = chkSdiv.Value
    itmX.SubItems(11) = chkBdiv.Value
    itmX.SubItems(12) = imnBnum.Value
    itmX.SubItems(13) = imnSnum.Value
    
'    '入力済みの場合は、前景色を変える
'    If itmX.SubItems(4) = INPUT_ON Or itmX.SubItems(7) <> "0" Then
'        varColor = DETAIL_FORECOLOR2
'    Else
'        varColor = DETAIL_FORECOLOR1
'    End If
'    itmX.ForeColor = varColor
'    For intIndex1 = 1 To MAX_COL
'        itmX.ListSubItems(intIndex1).ForeColor = varColor
'    Next intIndex1
    
    'リストビューをスクロールして、検出された ListItem を表示
    lsvMeisai.ListItems(lsvMeisai.ListItems.Count).EnsureVisible
    
    '行番号取得
    Call ListViewGetMaxRow
    
    ListViewSetItem = True
    
    Exit Function

ListViewSetItem_Err:

    Call MsgBox("リストビューへのデータ登録エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewSetItem_Err")

End Function

'目　的　　：リストビューからのデータ表示
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub ListViewGetItem()

    On Error GoTo ListViewGetItem_Err
    
    imnNo.Value = lsvMeisai.SelectedItem.Text
    cboIcode.Text = Trim(lsvMeisai.SelectedItem.SubItems(1))
    txtIname.Text = Trim(lsvMeisai.SelectedItem.SubItems(2))
    imnQty.Value = IIf(Trim(lsvMeisai.SelectedItem.SubItems(3)) <> "", lsvMeisai.SelectedItem.SubItems(3), 0)
    chkIdiv.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(4)), lsvMeisai.SelectedItem.SubItems(4), 0)

'    imnPrice1.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(5)), lsvMeisai.SelectedItem.SubItems(5), 0)
'    imnPrice2.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(6)), lsvMeisai.SelectedItem.SubItems(6), 0)
    imnPrice1.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(6)), lsvMeisai.SelectedItem.SubItems(6), 0)
    imnPrice2.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(5)), lsvMeisai.SelectedItem.SubItems(5), 0)
    
    imnPrice.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(7)), lsvMeisai.SelectedItem.SubItems(7), 0)
    cboBcode.Text = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(8)), lsvMeisai.SelectedItem.SubItems(8), "")
    lblBname.Caption = Trim(lsvMeisai.SelectedItem.SubItems(9))
    chkSdiv.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(10)), lsvMeisai.SelectedItem.SubItems(10), 0)
    chkBdiv.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(11)), lsvMeisai.SelectedItem.SubItems(11), 0)
    imnBnum.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(12)), lsvMeisai.SelectedItem.SubItems(12), 0)
    imnSnum.Value = IIf(IsNumeric(lsvMeisai.SelectedItem.SubItems(13)), lsvMeisai.SelectedItem.SubItems(13), 0)
            
    '金額がゼロ以外は注文分とみなす
    If imnPrice.Value <> 0 Then
        chkChumon.Value = 1
    End If
        
    Exit Sub
    
ListViewGetItem_Err:

   Call MsgBox("リストビューからデータ取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewGetItem_Err")

End Sub

'目　的　　：リストビューからのデータ削除
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function ListViewDelItem() As Boolean

    Dim itmX As ListItem
    Dim intPostion As Integer

    On Error GoTo ListViewDelItem_Err

    ListViewDelItem = False

    If MsgBox("明細を削除しますか？", vbYesNo + vbQuestion, "") = vbNo Then Exit Function
    
    '削除行の取得
    intPostion = imnNo.Value
    
    'リストビューのデータ検索（行番号が一致するデータがあったら削除）
    Set itmX = lsvMeisai.FindItem(intPostion, , , 0)
    If Not (itmX Is Nothing) Then
        'データ削除
        lsvMeisai.ListItems.Remove itmX.Index
        '行番号振り直し
        Call ListViewRefresh
    End If

    '行番号取得
    Call ListViewGetMaxRow

    ListViewDelItem = True

    Exit Function

ListViewDelItem_Err:

    Call MsgBox("リストビューからデータ削除エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewDelItem_Err")

End Function

Private Sub txtPnum_Validate(Cancel As Boolean)

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo txtPnum_Validate_Err

    If Trim(txtPnum.Text) = "" Then Exit Sub
    
    '受付データ
    strSQL = "SELECT * FROM DT010" & _
             " WHERE Odate = '" & lblOdate & "'" & _
             " AND Pnum = " & txtPnum.Text & _
             " ORDER BY Odate,Pnum"
    adoRecordset1.Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If adoRecordset1.EOF = True Then
        Cancel = True
        Call MsgBox("受付番号が存在しません。", vbOKOnly + vbCritical, "")
    End If
    adoRecordset1.Close
    
    Exit Sub
    
txtPnum_Validate_Err:
    
    Call MsgBox("受付番号フォーカス喪失時エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "txtPnum_Validate_Err")
    
End Sub

Private Sub txtSname_Click()

    If optSyori(0).Value = True Then
        '変更モードにする
        optSyori(1).Value = True
        txtOnum.Text = txtSname.Value
        Call txtOnum_Validate(False)
    End If

End Sub

Private Sub txtSname_DropDown()

    Call MaketxtSname(txtSname)

End Sub

Private Sub txtSname_GotFocus()

    txtSname.BackColor = FOCUS_STOP_COLOR
    Call SetImeMode(ActiveControl.hwnd, 4)
    
End Sub

Private Sub txtSname_LostFocus()

    txtSname.BackColor = FOCUS_NO_COLOR

End Sub

'2005/08/12 追加
Private Sub txtSname_Validate(Cancel As Boolean)
    
    If Trim(txtSname.Text) = "" Then Exit Sub
    If InStr(txtSname.Text, "注文分") <= 0 Then
        txtSname.Text = txtSname.Text & "(注文分)"
    End If

End Sub

Private Sub txtTel_GotFocus()

    txtTel.BackColor = FOCUS_STOP_COLOR

End Sub

Private Sub txtTel_LostFocus()

    txtTel.BackColor = FOCUS_NO_COLOR
    
End Sub

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub MakecboScode_Kana(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboScode_Kana_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        '得意先マスタ
        If Trim(strBuff1) = "" Then
            strSQL = "SELECT Bkana,Bname,Bcode FROM MT070" & _
                     " WHERE Fdiv = " & BUSINESS_DIV_EXHIBITION & " OR Fdiv = " & BUSINESS_DIV_ALL & _
                     " ORDER BY Bkana,Bname,Bcode"
        Else
            strSQL = "SELECT Bkana,Bname,Bcode FROM MT070" & _
                     " WHERE (Fdiv = " & BUSINESS_DIV_EXHIBITION & " OR Fdiv = " & BUSINESS_DIV_ALL & ")" & _
                     " AND Bkana LIKE '" & strBuff1 & "%'" & _
                     " ORDER BY Bkana,Bname,Bcode"
        End If
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Ctrl.AddItem .Fields("Bkana") & ";" & .Fields("Bname") & ";" & .Fields("Bcode")
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboScode_Kana_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboScode_Kana_Err")

End Sub

'目　的　　：合計の計算
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub Calc_Total()

    Dim intIndex1 As Integer
    Dim curTotal As Currency
    Dim varPrice As Variant

    On Error GoTo Calc_Total_Err
    
    curTotal = 0
    varPrice = 0
    For intIndex1 = 1 To lsvMeisai.ListItems.Count
        curTotal = curTotal + CCur(lsvMeisai.ListItems(intIndex1).SubItems(3))
        If Trim(lsvMeisai.ListItems(intIndex1).SubItems(7)) <> "" Then
            varPrice = varPrice + CDec(lsvMeisai.ListItems(intIndex1).SubItems(7))
        End If
    Next intIndex1
    
    imnQty_Total.Value = curTotal
    If varPrice > imnPrice_Total.MaxValue Then
        imnPrice_Total.Value = imnPrice_Total.MaxValue
    ElseIf varPrice < imnPrice_Total.MinValue Then
        imnPrice_Total.Value = imnPrice_Total.MinValue
    Else
        imnPrice_Total.Value = varPrice
    End If
    
    Exit Sub
    
Calc_Total_Err:

    Call MsgBox("合計の計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Total_Err")

End Sub

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
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
                     " WHERE Ikana LIKE '%" & strBuff1 & "%'" & _
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

'目　的　　：明細入力チェック
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function DoValidationChecks_Dst() As Boolean

    Dim strErrMsg As String
        
    On Error GoTo DoValidationChecks_Dst_Err
    
    If imnNo.Value > MAX_ROW Then
        strErrMsg = StrConv((MAX_ROW + 1), vbWide) & "行以上入力できません。"
        txtIname.SetFocus
        GoTo ErrorTrap:
    End If
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
    If imnPrice1.Value = 0 Then
        strErrMsg = "単価を入力してください。"
        If chkChumon.Value = 0 Then chkChumon.Value = 1
        DoEvents
        imnPrice1.SetFocus
        GoTo ErrorTrap:
    End If
    If Trim(cboBcode.Text) = "" Then
        strErrMsg = "買主を入力してください。"
        If chkChumon.Value = 0 Then chkChumon.Value = 1
        DoEvents
        cboBcode.SetFocus
        GoTo ErrorTrap:
    End If

    DoValidationChecks_Dst = True

    Exit Function
    
ErrorTrap:
    
    DoEvents
    DoValidationChecks_Dst = False
    Call MsgBox(strErrMsg & vbCrLf & Error$, vbOKOnly + vbCritical, "入力チェック")
    
    Exit Function
    
DoValidationChecks_Dst_Err:

    DoValidationChecks_Dst = False
    Call MsgBox("入力チェックエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "DoValidationChecks_Dst_Err")

End Function

'目　的　　：リストビューからの行番号取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function ListViewGetMaxRow() As Boolean

    On Error GoTo ListViewGetMaxRow_Err

    ListViewGetMaxRow = False

    '行番号取得
    imnNo.Value = lsvMeisai.ListItems.Count + 1

    ListViewGetMaxRow = True

    Exit Function

ListViewGetMaxRow_Err:

    Call MsgBox("リストビューからの行番号取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewGetMaxRow_Err")

End Function

'目　的　　：リストビューへのデータ挿入
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function ListViewInsItem() As Boolean
    
    Dim varRes As Variant
    Dim intPostion As Integer
    
    On Error GoTo ListViewInsItem_Err
    
    ListViewInsItem = False
    
    If lsvMeisai.ListItems.Count >= MAX_ROW Then
        Call MsgBox("これ以上明細を入力できません。", vbOKOnly + vbCritical, "")
        Exit Function
    End If
    
    varRes = InputBox("挿入する行番号を入力してください...", "", "")

    '入力値をチェック
    If Trim(varRes) = "" Then
        Call MsgBox("行番号を入力してください。", vbOKOnly + vbCritical, "")
        Exit Function
    End If
    If IsNumeric(varRes) = False Then
        Call MsgBox("行番号が不正です。", vbOKOnly + vbCritical, "")
        Exit Function
    End If

    If DoValidationChecks_Dst() = False Then Exit Function

    '編集行の取得
    intPostion = CInt(varRes)
    
    Call ListViewSetItem(intPostion, 1)

    '行番号振り直し
    Call ListViewRefresh

    ListViewInsItem = True

    Exit Function

ListViewInsItem_Err:

    Call MsgBox("リストビューへのデータ挿入エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewInsItem_Err")

End Function

'目　的　　：リストビューの行番号を振り直す
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function ListViewRefresh() As Boolean

    Dim intIndex1 As Integer

    On Error GoTo ListViewRefresh_Err

    ListViewRefresh = False

    lsvMeisai.Refresh
    For intIndex1 = 1 To lsvMeisai.ListItems.Count Step 1
        lsvMeisai.ListItems(intIndex1).Text = intIndex1
    Next intIndex1

    ListViewRefresh = True

    Exit Function

ListViewRefresh_Err:

    Call MsgBox("リストビューの行番号を振り直しエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ListViewRefresh_Err")

End Function

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Sub MakecboScode(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MakecboScode_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        '得意先マスタ
        strSQL = "SELECT Bkana,Bname,Bcode FROM MT070" & _
                 " WHERE Fdiv = " & BUSINESS_DIV_EXHIBITION & " OR Fdiv = " & BUSINESS_DIV_ALL & _
                 " ORDER BY Bcode"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            Ctrl.AddItem .Fields("Bcode") & ";" & .Fields("Bname")
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MakecboScode_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MakecboScode_Err")

End Sub

'目　的　　：売立金額の計算
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／１８
'更新履歴　：
'
Private Function Calc_Price() As Boolean

    Dim varPrice As Variant

    On Error GoTo Calc_Price_Err
    
    Calc_Price = False
    
    varPrice = CDec(imnQty.Value) * CDec(imnPrice1.Value)
    If varPrice > imnPrice.MaxValue Then
        Call MsgBox("売立金額が大きすぎます。", vbOKOnly + vbCritical, "")
        DoEvents
    ElseIf varPrice < imnPrice.MinValue Then
        Call MsgBox("売立金額が小さすぎます。", vbOKOnly + vbCritical, "")
        DoEvents
    Else
        imnPrice.Value = CDec(imnQty.Value) * CDec(imnPrice1.Value)
        Calc_Price = True
    End If
    
    Exit Function
    
Calc_Price_Err:

    Call MsgBox("売立金額の計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Calc_Price_Err")

End Function

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２１
'更新履歴　：
'
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
                If .Fields("Sdate") <= Trim(lblOdate.Caption) And Trim(lblOdate.Caption) <= .Fields("Fdate") Then
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

'目　的　　：コンボボックスの作成
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００５／０８／１９
'更新履歴　：
'
Private Sub MaketxtSname(Ctrl As Control)

    Dim strBuff1 As String
    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo MaketxtSname_Err
    
    Screen.MousePointer = vbHourglass
    
    strBuff1 = Trim(Ctrl.Text)
    Ctrl.Clear
    
    With adoRecordset1
        strSQL = "SELECT Onum,Sname FROM DT030" & _
                 " WHERE Odate = '" & lblOdate.Caption & "'" & _
                 " ORDER BY Onum,Sname"
        .Open strSQL, g_clsAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            If Not IsNull(.Fields("Sname")) Then
                Ctrl.AddItem .Fields("Onum") & ";" & .Fields("Sname")
            Else
                Ctrl.AddItem .Fields("Onum") & ";" & ""
            End If
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Ctrl.Text = strBuff1
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
MaketxtSname_Err:

    Screen.MousePointer = vbDefault
    Call MsgBox("コンボボックス作成エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "MaketxtSname_Err")

End Sub

