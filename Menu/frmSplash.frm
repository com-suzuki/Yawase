VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   ClientHeight    =   4560
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7320
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BorderStyle     =   0  'Ç»Çµ
      Height          =   4590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1425
         Left            =   420
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   1365
         ScaleWidth      =   1830
         TabIndex        =   2
         Top             =   660
         Width           =   1890
      End
      Begin VB.Label lblUrl 
         Caption         =   "http://www.com-e.co.jp"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4740
         TabIndex        =   10
         Tag             =   "Company"
         Top             =   3600
         Width           =   2115
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'âEëµÇ¶
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Tag             =   "LicenseTo"
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   19.5
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   2520
         TabIndex        =   9
         Tag             =   "Product"
         Top             =   1260
         Width           =   4488
      End
      Begin VB.Label lblCompanyProduct 
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2505
         TabIndex        =   8
         Tag             =   "CompanyProduct"
         Top             =   765
         Width           =   4425
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'âEëµÇ¶
         AutoSize        =   -1  'True
         Caption         =   "for Windows 2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4305
         TabIndex        =   7
         Tag             =   "Platform"
         Top             =   2400
         Width           =   2700
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'âEëµÇ¶
         AutoSize        =   -1  'True
         Caption         =   "Version 1.00"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5424
         TabIndex        =   6
         Tag             =   "Version"
         Top             =   2760
         Width           =   1584
      End
      Begin VB.Label lblWarning 
         Caption         =   "Ç±ÇÃêªïiÇÕì˙ñ{çëíòçÏå†ñ@Ç®ÇÊÇ—çëç€èñÒÇ…ÇÊÇËï€åÏÇ≥ÇÍÇƒÇ¢Ç‹Ç∑ÅB"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Tag             =   "Warning"
         Top             =   4260
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         Caption         =   "áä ∫—•¥›ºﬁ∆±ÿ›∏ﬁ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4710
         TabIndex        =   5
         Tag             =   "Company"
         Top             =   3330
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright(C) 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4710
         TabIndex        =   4
         Tag             =   "Copyright"
         Top             =   3120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    lblProductName.Caption = SYSTEM_NAME

End Sub

