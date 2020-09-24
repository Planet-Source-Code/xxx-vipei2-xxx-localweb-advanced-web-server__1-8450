VERSION 5.00
Begin VB.Form frmMIME 
   Caption         =   "Associate MIME Type"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   Icon            =   "frmMIME.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   4215
      Begin VB.Label Label1 
         Caption         =   $"frmMIME.frx":0442
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove Header"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Header..."
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox MIMEList 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmMIME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    frmAddMIME.Show 1
End Sub

Private Sub Command2_Click()
Dim Response
    
    Response = MsgBox("Are you sure you want to remove that MIME Header?", vbQuestion + vbYesNo, "LocalWEB")
    If Response = vbYes Then
      MimeList.RemoveItem (MimeList.ListIndex)
    End If

End Sub

Private Sub Command3_Click()
Dim FileHandle
Dim f

    FileHandle = FreeFile
    If FileExists("mime.lst") Then
      Kill "mime.lst"
    End If
    
    Open "mime.lst" For Output As #FileHandle
    For f = 0 To MimeList.ListCount - 1
      Write #FileHandle, MimeList.List(f)
    Next f
    Close #FileHandle
    
    Unload Me
    
End Sub

Private Sub Command4_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim FileHandle
Dim Header As String

    If FileExists("mime.lst") Then
      FileHandle = FreeFile
      Open "mime.lst" For Input As #FileHandle
      While Not EOF(FileHandle)
        Input #FileHandle, Header
        MimeList.AddItem Header
      Wend
      Close #FileHandle
    End If
        
    
End Sub
