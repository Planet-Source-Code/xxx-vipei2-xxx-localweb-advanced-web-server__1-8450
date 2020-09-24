VERSION 5.00
Begin VB.Form frmTelnet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Telnet Server Configuration"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmTelnet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame TNetFrame 
      Caption         =   " Settings "
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
      Begin VB.Frame TNetFrame2 
         Caption         =   " Telnet Server Password "
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   3975
         Begin VB.TextBox TelnetPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.CheckBox TelnetEnabled 
         Caption         =   "Enable LocalWEB's Telnet Server"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label TNetText 
      Caption         =   $"frmTelnet.frx":0442
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image TNetImage 
      Height          =   480
      Left            =   120
      Picture         =   "frmTelnet.frx":0548
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slValue As String
Dim slData As String
Dim rc


Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub Enabled_Click()

End Sub

Private Sub Command2_Click()

    slValue = "TelnetEnabled"
    slData = TelnetEnabled.Value
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "TelnetPassword"
    slData = UCase(TelnetPassword.Text)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    MsgBox "The Telnet Server settings have been saved.", vbInformation + vbOKOnly, "LocalWEB Telnet Configuration"
    Unload Me

End Sub

Private Sub Form_Load()

    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\LocalWEB\Settings"
         
    TelnetEnabled.Value = GetRegValue(hKey, SubKey, "TelnetEnabled", "1")
    TelnetPassword = GetRegValue(hKey, SubKey, "TelnetPassword", "PASSWORD")
    
    'check if telnet server is enabled
    
    If TelnetEnabled.Value = 0 Then
      TelnetPassword.Enabled = False
    Else
      TelnetPassword.Enabled = True
    End If
    
End Sub

Private Sub TelnetEnabled_Click()

    If TelnetEnabled.Value = 1 Then
      TelnetPassword.Enabled = True
    Else
      TelnetPassword.Enabled = False
    End If
End Sub
