VERSION 5.00
Begin VB.Form frmIPAddress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add IP Address"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "frmIPAddress.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Num4 
      Height          =   285
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Num3 
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Num2 
      Height          =   285
      Left            =   960
      MaxLength       =   3
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Num1 
      Height          =   285
      Left            =   240
      MaxLength       =   3
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address To Exclude"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmIPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim IPAddress As String
Dim i
Dim itm As ListItem

    IPAddress = Num1 + "." + Num2 + "." + Num3 + "." + Num4
   ' For i = 0 To frmWizard.ExceptionList.ListCount - 1
   '   If frmWizard.ExceptionList.List(i) = IPAddress Then
   '
   '     MsgBox "The IP Address " + IPAddress + " is already in the Exception List.", vbOKOnly + vbExclamation, "LocalWEB Warning"
   '     Num1.Text = ""
   '     Num2.Text = ""
   '     Num3.Text = ""
   '     Num4.Text = ""
   '     Num1.SetFocus
   '     Exit Sub
   '   End If
   ' Next i
    Set itm = frmServerProperties.ExceptionList.ListItems.Add(, , IPAddress, , 3)
    'frmWizard.ExceptionList.AddItem IPAddress
    Unload Me
    
End Sub

Private Sub Command2_Click()

    Unload Me
    
End Sub
