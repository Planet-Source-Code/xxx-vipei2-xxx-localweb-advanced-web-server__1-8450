VERSION 5.00
Begin VB.Form frmDocuments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Default Documents"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmDocuments.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FooterFrame 
      Caption         =   "Page Footer"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   5175
      Begin VB.TextBox FooterCode 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   4935
      End
      Begin VB.CheckBox FooterEnabled 
         Caption         =   "Enable Page Footer"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Label FooterHTML 
         Caption         =   "Footer HTML"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame DocFrane 
      Caption         =   "Default Document List"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5175
      Begin VB.CommandButton RemoveDocument 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton AddDocument 
         Caption         =   "&Add..."
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.ListBox DefaultList 
         Height          =   1425
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton MoveDown 
         Caption         =   "È"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Move down list"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton MoveUp 
         Caption         =   "Ç"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Move up list"
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Label DocDescription 
      Caption         =   $"frmDocuments.frx":0442
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image DocImage 
      Height          =   480
      Left            =   120
      Picture         =   "frmDocuments.frx":051A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slValue As String
Dim slData As String
Dim rc
Private Sub Command1_Click()
Dim Store As String

    'check to see if on first item
    'if so don't try and move any items
    
    If DefaultList.ListIndex = 0 Then
      Exit Sub
    Else
      Store = DefaultList.List(DefaultList.ListIndex - 1)
      DefaultList.List(DefaultList.ListIndex - 1) = DefaultList.List(DefaultList.ListIndex)
      DefaultList.List(DefaultList.ListIndex) = Store
    End If
    
    
End Sub

Private Sub Command2_Click()
Dim Store As String

    'are we at bottom of list and wanting to move an item down?
    'If so don't move item
    
    If DefaultList.ListIndex = DefaultList.ListCount - 1 Then
      Exit Sub
    Else
      Store = DefaultList.List(DefaultList.ListIndex + 1)
      DefaultList.List(DefaultList.ListIndex + 1) = DefaultList.List(DefaultList.ListIndex)
      DefaultList.List(DefaultList.ListIndex) = Store
    End If
    
End Sub

Private Sub Command3_Click()

    frmDocumentsAdd.Show 1
End Sub

Private Sub Command4_Click()
Dim Response
    
    Response = MsgBox("Are you sure you want to remove that Default Document?", vbQuestion + vbYesNo, "LocalWEB")
    If Response = vbYes Then
      DefaultList.RemoveItem (DefaultList.ListIndex)
    End If
End Sub

Private Sub AddDocument_Click()

    frmDocumentsAdd.Show 1
End Sub

Private Sub Command5_Click()
Dim FileHandle
Dim f
    
    If FileExists(App.Path & "\default.lst") Then
      Kill App.Path & "\default.lst"
    End If
     
    FileHandle = FreeFile
    Open App.Path & "\default.lst" For Output As #FileHandle
    For f = 0 To DefaultList.ListCount - 1
      Write #FileHandle, DefaultList.List(f)
    Next f
    Close #FileHandle
    
    'Write footer info to registry
    
    slValue = "FooterEnabled"
    slData = FooterEnabled.Value
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "FooterCode"
    slData = FooterCode.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    Unload Me
    
    
End Sub

Private Sub Command6_Click()

    Unload Me
    
End Sub

Private Sub FooterEnabled_Click()

    If FooterEnabled.Value = 1 Then
      FooterCode.Enabled = True
    Else
      FooterCode.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()
Dim FileHandle
Dim FName As String
 
    'Get Page Footer Information
    
    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\LocalWEB\Settings"
    FooterEnabled.Value = GetRegValue(hKey, SubKey, "FooterEnabled", "1")
    FooterCode = GetRegValue(hKey, SubKey, "FooterCode", "<p><font color=&quot;#0000FF&quot; size=&quot;1&quot;>This site is served by LocalWEB</font></p>")
    
    If FooterEnabled.Value = 1 Then
      FooterCode.Enabled = True
    Else
      FooterCode.Enabled = True
    End If
    
      FileHandle = FreeFile
      Open App.Path & "\default.lst" For Input As #FileHandle
      While Not EOF(FileHandle)
        Input #FileHandle, FName
        DefaultList.AddItem FName
      Wend
      Close #FileHandle

End Sub

