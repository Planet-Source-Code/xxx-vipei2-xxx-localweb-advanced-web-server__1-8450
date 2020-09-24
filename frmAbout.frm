VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About LocalWEB"
   ClientHeight    =   5700
   ClientLeft      =   3165
   ClientTop       =   2565
   ClientWidth     =   5910
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Trademarks"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Project Contributors"
      Height          =   1695
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   5415
      Begin VB.ListBox ContList 
         Height          =   1035
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label4 
         Caption         =   "Thanks to the following for help and encouragement:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "User Interface Design and Programming By Jeremy Arnold"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Original Design and Programming By Phil Curnow"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   $"frmAbout.frx":0742
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Width           =   4095
   End
   Begin VB.Label VerInfo 
      Caption         =   "Label4"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "http://www.west-street.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      MouseIcon       =   "frmAbout.frx":0881
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1560
      Width           =   2085
   End
   Begin VB.Label Label2 
      Caption         =   "A Freeware 32 bit Web Server for Windows 9x and NT and Windows 2000"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "LocalWEB Web Server Release Version"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MajorVersion = 1
Private Const MinorVersion = 0
Private Const Build = 12

'Added by STF
'ShowURL private declarations

Private Const SW_SHOW = 5
Private Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
   String, ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias _
   "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
   String, ByVal lpResult As String) As Long


Sub ShowURL(PageURL As String)
         Dim FileName, Dummy As String
         Dim BrowserExec As String * 255
         Dim RetVal As Long
         Dim FileNumber As Integer
   
         ' First, create a known, temporary HTML file
         BrowserExec = Space(255)
         'FileName = "c:\temphtm.HTM"
         FileName = "temphtm.HTM"
         FileNumber = FreeFile                    ' Get unused    file Number
         Open FileName For Output As #FileNumber  ' Create temp    HTML file
             Write #FileNumber, "<HTML> <\HTML>"  ' Output text
         Close #FileNumber                        ' Close file
         ' Then find the application associated with it
         RetVal = FindExecutable(FileName, Dummy, BrowserExec)
         BrowserExec = Trim(BrowserExec)
         ' If an application is found, launch it!
         If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
             MsgBox "Could not find associated Browser", vbExclamation, _
               "Browser Not Found"
         Else
             RetVal = ShellExecute(Me.hwnd, "open", BrowserExec, _
               PageURL, Dummy, SW_SHOWNORMAL)
             If RetVal <= 32 Then        ' Error
                 MsgBox "Web Page not Opened", vbExclamation, "URL", "Failed" ' ""
             End If
         End If
         Kill FileName                   ' delete temp HTML file

End Sub


Private Sub Command1_Click()

    Unload Me
End Sub

Private Sub Command2_Click()

    frmTradeMarks.Show 1
End Sub

Private Sub Form_Load()

    VerInfo.Caption = App.Major & "." & App.Minor & "." & App.Revision
    
    'Populate Contributors List
    
    ContList.AddItem "Steven T. Fricke"
    ContList.AddItem "Gerard Gilbert"
    ContList.AddItem "Bogdan Florin"
    ContList.AddItem "Pinhas Ifergan - DNS4EVER.COM"
End Sub

Private Sub Label10_Click()
    ShowURL Trim(Label10.Caption)
End Sub


