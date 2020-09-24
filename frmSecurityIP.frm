VERSION 5.00
Begin VB.Form frmSecurityIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add IP Address"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frmSecurityIP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Num1 
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Num2 
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Num3 
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Num4 
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address To Block"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSecurityIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim IPAddress As String
Dim i
Dim itm As ListItem

    IPAddress = Num1 + "." + Num2 + "." + Num3 + "." + Num4
    Set itm = frmServerProperties.ServerDeny.ListItems.Add(, , IPAddress, , 3)
    
    Unload Me
End Sub

Private Sub Command2_Click()

    Unload Me
    
End Sub

