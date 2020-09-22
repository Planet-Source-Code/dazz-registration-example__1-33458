VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Example"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   6330
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar s 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3705
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   8546
            TextSave        =   "9:30 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   4560
      Left            =   0
      Picture         =   "fMain.frx":0442
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   6360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuReg 
         Caption         =   "Register"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'for some reason it works better done twice =/
    frmReg.Rego
    frmReg.Rego
    S.Panels(1).Text = frmReg.txtRegName
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Created by Dazz" & vbNewLine & _
            "KeyGen Mod by sumone else but modified!" & vbNewLine & _
            "Registered To: " & frmReg.txtRegName.Text & vbNewLine & _
            "Serial No: " & frmReg.lbSerialNum.Text
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuReg_Click()
    frmReg.Show
End Sub

