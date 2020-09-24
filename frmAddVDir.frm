VERSION 5.00
Begin VB.Form frmAddVDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Virtual Directory"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmAddVDir.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox PathName 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox VirtualDir 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAddVDir.frx":0442
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Virtual Directory Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmAddVDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    frmBrowseDirs.Show 1

End Sub

Private Sub Command2_Click()
Dim itm As ListItem

    If Right(VirtualDir, 1) <> "/" Then
      VirtualDir = VirtualDir & "/"
    End If
    Set itm = frmServerProperties.VirtualDirectories.ListItems.Add(, , VirtualDir)
    itm.SubItems(1) = PathName
   
    Unload Me
    
    
End Sub

Private Sub Command3_Click()
Dim itm As ListItem

    If frmAddVDir.Caption = "Edit Virtual Directory" Then
      If Right(VirtualDir, 1) <> "/" Then
        VirtualDir = VirtualDir & "/"
      End If
      Set itm = frmServerProperties.VirtualDirectories.ListItems.Add(, , VirtualDir)
      itm.SubItems(1) = PathName
    End If
    
    Unload Me
    
End Sub

