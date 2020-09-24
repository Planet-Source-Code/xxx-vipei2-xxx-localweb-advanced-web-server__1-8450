VERSION 5.00
Begin VB.Form frmAddMIME 
   Caption         =   "Add MIME Type"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmAddMIME.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox MIMEHeader 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox FileExt 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Content Type (eg image/jpeg):"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "File Extension (eg JPG):"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmAddMIME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Len(FileExt) = 0 Or Len(MIMEHeader) = 0 Then
      MsgBox "You must supply a file extension AND a MIME Header.", vbExclamation + vbOKOnly, "LocalWEB"
      Exit Sub
    End If
    frmServerProperties.MimeList.AddItem FileExt + "," + MIMEHeader
    FileExt = ""
    MIMEHeader = ""
    FileExt.SetFocus
    Unload Me
    
End Sub

Private Sub Command2_Click()

    Unload Me
    
End Sub
