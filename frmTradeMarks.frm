VERSION 5.00
Begin VB.Form frmTradeMarks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trade Marks"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "frmTradeMarks.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox MarkList 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   $"frmTradeMarks.frx":0442
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "All trademarks are acknowledged.  The following trademarks are registered and the property of the respective companies."
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "® "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmTradeMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    MarkList.AddItem "Windows 95, Windows 98, Windows NT and Windows 2000 is a Trademark of Microsoft Corp."
    MarkList.AddItem "Visual Basic is a Trademark of Microsoft Corp."
    
End Sub
