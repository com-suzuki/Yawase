VERSION 5.00
Object = "{8238B41D-9CEC-11D4-AC6B-00004CF3B072}#1.0#0"; "COMFNPG32.ocx"
Begin VB.Form frmCount 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   ClientHeight    =   465
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5100
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin FineProgress.fpProgressBar fpProgressBar1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      BarColor        =   8388608
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   100
      Style           =   0
      Unit            =   "%"
      Value           =   0
      ValueShow       =   -1  'True
   End
   Begin VB.TextBox txtDummy 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "íÜÅ@é~"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_blnCancel As Boolean

Private Sub cmdCancel_Click()
    
    g_blnCancel = True

End Sub

Private Sub Form_Load()
    
    g_blnCancel = False

End Sub


