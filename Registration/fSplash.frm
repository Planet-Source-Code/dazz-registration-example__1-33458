VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   240
         Top             =   600
      End
      Begin VB.Label Label2 
         Caption         =   "The Registered Version Doesnt Display this Nag/Delay Screen =P"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   3120
         Width           =   4695
      End
      Begin VB.Label R1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "RRP $19.95"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgLogo 
         Height          =   945
         Left            =   360
         Picture         =   "fSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 2002 the f00ls!"
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
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblCompany 
         Caption         =   "The Nagging Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning...Dont pay me and ill piss you off!NAGNAGNAG!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1.0.0[Shareware]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4590
         TabIndex        =   4
         Top             =   2700
         Width           =   2265
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform"
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
         Left            =   4680
         TabIndex        =   5
         Top             =   2340
         Width           =   2115
      End
      Begin VB.Label lblProductName 
         BackStyle       =   0  'Transparent
         Caption         =   "Best Program in the world! "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         TabIndex        =   7
         Top             =   1200
         Width           =   5700
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Pain in Da Ass!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   2040
         TabIndex        =   6
         Top             =   705
         Width           =   2610
      End
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim s As Integer


Private Sub Form_Load()
    frmReg.Rego
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim lod As String
s = s + 1
    R1.Caption = lod
    pb1.Value = s
        If s = 100 Then
            Timer1.Enabled = False
            fMain.Show
            Unload Me
            MsgBox "Dont forget to buy my software!!!111!", vbInformation

        End If
End Sub
