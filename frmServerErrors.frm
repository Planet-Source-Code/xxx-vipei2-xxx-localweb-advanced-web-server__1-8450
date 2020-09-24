VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmServerErrors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Error Messages"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmServerErrors.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox CGIServerBarred 
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox CGIUserBarred 
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox CGINotFound 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Error404 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "CGI Barred On This Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "CGI Barred On Your Machine:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "CGI Script Not Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "404 Not Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label ErrDesc 
      Caption         =   $"frmServerErrors.frx":0442
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmServerErrors.frx":0514
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmServerErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slValue As String
Dim slData As String
Dim rc

Private Sub Command1_Click()
    
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*|HTML Files(*.htm)|*.htm|HTML Files(*.html)|*.html"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    Error404.Text = CommonDialog1.FileName
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

    

    
End Sub

Private Sub Command2_Click()
 
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*|HTML Files(*.htm)|*.htm|HTML Files(*.html)|*.html"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    CGINotFound.Text = CommonDialog1.FileName
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub Command3_Click()
 
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*|HTML Files(*.htm)|*.htm|HTML Files(*.html)|*.html"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    CGIUserBarred.Text = CommonDialog1.FileName
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub Command4_Click()
 
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*|HTML Files(*.htm)|*.htm|HTML Files(*.html)|*.html"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    CGIServerBarred.Text = CommonDialog1.FileName
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub Command5_Click()

    slValue = "Error404"
    slData = Error404.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "CGINotFound"
    slData = CGINotFound.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "CGIUserBarred"
    slData = CGIUserBarred.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "CGIServerBarred"
    slData = CGIServerBarred.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    Unload Me
    
End Sub

Private Sub Command6_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\LocalWEB\Settings"
         
    Error404.Text = GetRegValue(hKey, SubKey, "Error404", "Default")
    CGINotFound.Text = GetRegValue(hKey, SubKey, "CGINotFound", "Default")
    CGIUserBarred.Text = GetRegValue(hKey, SubKey, "CGIUserBarred", "Default")
    CGIServerBarred.Text = GetRegValue(hKey, SubKey, "CGIServerBarred", "Default")
    
    
End Sub
