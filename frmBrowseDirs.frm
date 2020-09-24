VERSION 5.00
Begin VB.Form frmBrowseDirs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse Directories"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmBrowseDirs.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   3975
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmBrowseDirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    frmAddVDir.PathName = Dir1.Path
    Unload Me
    
End Sub

Private Sub Command2_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    If frmAddVDir.Caption = "Edit Virtual Directory" Then
      Dir1.Path = frmAddVDir.PathName
    End If
    
End Sub
