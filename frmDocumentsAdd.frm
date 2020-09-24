VERSION 5.00
Begin VB.Form frmDocumentsAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add a Default Document"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmDocumentsAdd.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox DocName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Document Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmDocumentsAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Len(DocName) = 0 Then
      MsgBox "You must supply a document name.", vbExclamation + vbOKOnly, "LocalWEB"
      Exit Sub
    End If
    frmServerProperties.DefaultList.AddItem DocName
    Unload Me
    
End Sub
