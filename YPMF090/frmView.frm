VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmView 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "éÛïtñæç◊"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10575
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmView.frx":0CFA
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   13361
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Pnum"
         Caption         =   "éÛïtî‘çÜ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Line"
         Caption         =   "çs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Iname"
         Caption         =   "êAñÿñºèÃ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Price"
         Caption         =   "îÑóßã‡äz"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Bcode"
         Caption         =   "îÉéÂ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Idiv"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3734.929
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   945.071
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "ñﬂÅ@ÇÈ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Top             =   7800
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   120
      Top             =   7860
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=YAWASESRC"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "YAWASESRC"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   $"frmView.frx":0D0F
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_intPnum As Integer

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Dim strSQL As String
    
    On Error GoTo Form_Load_Err

    'éÛïtÉfÅ[É^
    strSQL = "SELECT DT011.Odate, DT011.Pnum, DT011.Line, DT011.Iname, DT011.Qty, DT021.Price, DT021.Bcode," & _
             " (CASE DT021.Idiv WHEN 1 THEN 'ÉÑÉÅ' ELSE '' END) AS Idiv" & _
             " FROM DT021 INNER JOIN" & _
             " DT020 ON DT021.Ocode = DT020.Ocode RIGHT OUTER JOIN" & _
             " DT011 ON DT021.Pnum = DT011.Pnum AND DT021.PnumLine = DT011.Line AND DT020.Odate = DT011.Odate" & _
             " WHERE DT011.Odate = '" & g_strOdate & "'" & _
             " AND DT011.Pnum = " & m_intPnum & _
             " ORDER BY DT011.Line"
    Adodc1.RecordSource = strSQL
    Adodc1.Refresh
    
    Exit Sub

Form_Load_Err:

    Call MsgBox("ÉtÉHÅ[ÉÄÉçÅ[ÉhéûÉGÉâÅ[ÅIÅI" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Form_Load_Err")

End Sub
