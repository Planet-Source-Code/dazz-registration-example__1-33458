VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register..."
   ClientHeight    =   2205
   ClientLeft      =   3885
   ClientTop       =   3510
   ClientWidth     =   3645
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReg.frx":08CA
   ScaleHeight     =   2205
   ScaleWidth      =   3645
   Begin VB.TextBox lbSerialNum 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "^"
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtRegName 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "."
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox l1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtserial 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Serial"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim keyz As String
Dim serial As String
Dim Correct As Boolean
Dim kGen As String
Private Sub cmdOK_Click()

    If txtserial.Text = "" Then
        MsgBox "Enter a serial number!"
        Exit Sub
    Else
        SaveSettingString HKEY_CURRENT_USER, "Software\TestProgram\License\", "Name", txtName.Text
        SaveSettingString HKEY_CURRENT_USER, "Software\TestProgram\License\", "Serial", txtserial.Text
        MsgBox "The program will now end to see if serial is valid, you can reopen it again after this!"
        End
        
    End If
End Sub

Private Sub txtName_Change()
    keyz = KeyGen(txtName, "Test", 2)
    l1 = keyz
End Sub
Public Sub Rego()
On Error GoTo e:
Dim username, Stored As String
    'get the stored settings
    username = GetSettingString(HKEY_CURRENT_USER, "Software\TestProgram\License\", "Name", "")
    Stored = GetSettingString(HKEY_CURRENT_USER, "Software\TestProgram\License\", "Serial", "")
    'use the stored settings!
    txtRegName = username
    lbSerialNum = Stored
    'create the "TRUE KEY!"
    kGen = KeyGen(txtRegName, "Test", 2)

        'set the rego to false
        Correct = False
        
            'check the serial to see if its good!
            If lbSerialNum = kGen Then
                Correct = True 'if its good to to true
            Else
                Correct = False 'if its bad set to false
            End If
                If Correct = False Then
                    fSplash.Timer1.Enabled = True
                    fSplash.Show
                    fMain.mnuReg.Visible = True
                    fMain.Caption = "Registration Example-[Unregistered Version]"
                    lbSerialNum.Text = "Un-Registered"
                    txtRegName = "Un-Registered"
                ElseIf Correct = True Then
                    fMain.Show
                    fSplash.Visible = False
                    fSplash.Timer1.Enabled = False
                    fMain.mnuReg.Visible = False
                End If
'for all error handling!
e:
    If Err.Number <> 0 Then
                MsgBox Err.Description, vbCritical
                
            Exit Sub
        End If
End Sub

