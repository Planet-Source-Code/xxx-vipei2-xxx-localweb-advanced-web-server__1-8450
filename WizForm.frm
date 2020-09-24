VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LocalWEB Configuration Wizard"
   ClientHeight    =   5745
   ClientLeft      =   2700
   ClientTop       =   2100
   ClientWidth     =   8370
   Icon            =   "WizForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CmDlg 
      Left            =   1800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      Height          =   135
      Left            =   120
      TabIndex        =   34
      Top             =   4800
      Width           =   8175
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Index           =   4
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton Command8 
         Caption         =   "&Remove Type"
         Height          =   375
         Left            =   3120
         TabIndex        =   40
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Add Type..."
         Height          =   375
         Left            =   3120
         TabIndex        =   39
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox MimeList 
         Height          =   2400
         Left            =   240
         TabIndex        =   38
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   $"WizForm.frx":0442
         Height          =   855
         Left            =   840
         TabIndex        =   28
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   240
         Picture         =   "WizForm.frx":04F9
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Index           =   3
      Left            =   600
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   4695
         Begin VB.CheckBox AllowCGI 
            Caption         =   "Allow CGI Scripts To Be Run"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label8 
            Caption         =   "If you want LocalWEB to stop execution of all CGI scripts, simply un-check the box above."
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   4455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Except From The Following IP Addresses"
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   4695
         Begin VB.CommandButton Command7 
            Caption         =   "&Remove"
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "&Add..."
            Height          =   255
            Left            =   3240
            TabIndex        =   13
            Top             =   960
            Width           =   1095
         End
         Begin VB.ListBox ExceptionList 
            Height          =   645
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label16 
            Caption         =   "The following machines with these IP Addresses will not be able to execute CGI Scripts."
            Height          =   495
            Left            =   840
            TabIndex        =   41
            Top             =   360
            Width           =   3615
         End
         Begin VB.Image Image7 
            Height          =   480
            Left            =   120
            Picture         =   "WizForm.frx":093B
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Label Label7 
         Caption         =   $"WizForm.frx":0D7D
         Height          =   615
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   3735
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   240
         Picture         =   "WizForm.frx":0E11
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame6 
         Caption         =   " Enable/Disable Logging "
         Height          =   2535
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   4695
         Begin VB.ComboBox LogFormat 
            Height          =   315
            Left            =   1320
            TabIndex        =   35
            Top             =   1200
            Width           =   2895
         End
         Begin VB.CheckBox EnableLogging 
            Caption         =   "Enable Event Logging"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Select the format you wish the logfile to be written in.  The default is the LocalWEB Format."
            Height          =   495
            Left            =   240
            TabIndex        =   37
            Top             =   1800
            Width           =   4095
         End
         Begin VB.Label Label13 
            Caption         =   "Logfile Format:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Check the box if you require event logging."
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   3135
         End
      End
      Begin VB.Label Label3 
         Caption         =   $"WizForm.frx":111B
         Height          =   855
         Left            =   840
         TabIndex        =   20
         Top             =   360
         Width           =   3735
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   240
         Picture         =   "WizForm.frx":11D8
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Finish 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Back 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Next 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Index           =   0
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "You DO NOT have to close and re-run LocalWEB for any changes to take effect."
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   33
         Top             =   3600
         Width           =   4335
      End
      Begin VB.Label Label10 
         Caption         =   $"WizForm.frx":161A
         Height          =   1095
         Left            =   360
         TabIndex        =   32
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Configuration Wizard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"WizForm.frx":1746
         Height          =   1455
         Left            =   360
         TabIndex        =   30
         Top             =   720
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   4695
         Begin VB.TextBox LogFile 
            Height          =   285
            Left            =   840
            TabIndex        =   19
            Top             =   960
            Width           =   3555
         End
         Begin VB.Label Label2 
            Caption         =   "Supply the location of where you would like the log file generated by LocalWEB to be saved. For Example C:\WWW\LOGS"
            Height          =   615
            Left            =   840
            TabIndex        =   18
            Top             =   240
            Width           =   3615
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   120
            Picture         =   "WizForm.frx":18B5
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4695
         Begin VB.TextBox HomePage 
            Height          =   285
            Left            =   840
            TabIndex        =   9
            Top             =   1200
            Width           =   3435
         End
         Begin VB.Label Label1 
            Caption         =   "Supply the location of where the pages that LocalWEB is serving can be found on this machine. For example: C:\WWW\HOMEPAGE"
            Height          =   615
            Left            =   720
            TabIndex        =   16
            Top             =   240
            Width           =   3735
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "WizForm.frx":1CF7
            Top             =   240
            Width           =   480
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Welcome to the LocalWEB Configuration Wizard.  "
      Height          =   615
      Left            =   240
      TabIndex        =   29
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Image Image6 
      Height          =   2595
      Left            =   360
      Picture         =   "WizForm.frx":2139
      Top             =   480
      Width           =   2520
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StepCount As Integer
Dim LastStep As Integer


   
Private Sub Back_Click()
 If StepCount = 0 Then
      Exit Sub
    Else
      Frame1(StepCount).Visible = False
      StepCount = StepCount - 1
      DisplayStep StepCount, LastStep
    End If
End Sub

Private Sub cmdBrowse1_Click()

End Sub

Private Sub Command1_Click()
Dim phkResult As Long
Dim slValue As String
Dim slData As String
Dim MyString As String
Dim rc


    
    
    hKey = HKEY_LOCAL_MACHINE
    
    'set registry subkey
    SubKey = "SOFTWARE\LocalWEB\Settings"
    slValue = "HomePage"
    slData = "c:\www\root\"
    
    
    
    'If CreateRegKey(SubKey) Then
    '  rc = SetRegValue(hKey, SubKey, slValue, slData)
    'Else
    '  MsgBox "Cannot create registry value"
    'End If
    
    
End Sub





Private Sub Command4_Click()

    frmAddMIME.Show 1
End Sub

Private Sub Command5_Click()

    Unload Me
    
End Sub

Private Sub Command6_Click()

    frmIPAddress.Show 1
    
End Sub

Private Sub Command7_Click()
Dim Response
    
    Response = MsgBox("Are you sure you want to allow the machine with IP Address " + ExceptionList.List(ListIndex) + " to run CGI Scripts?", vbQuestion + vbYesNo, "LocalWEB")
    If Response = vbYes Then
      ExceptionList.RemoveItem (ExceptionList.ListIndex)
    End If
End Sub

Private Sub Command8_Click()
Dim Response
    
    Response = MsgBox("Are you sure you want to remove that MIME Header?", vbQuestion + vbYesNo, "LocalWEB")
    If Response = vbYes Then
      MIMEList.RemoveItem (MIMEList.ListIndex)
    End If
End Sub

Private Sub EnableLogging_Click()

    If EnableLogging.Value = 0 Then
      LogFormat.Enabled = False
    Else
      LogFormat.Enabled = True
    End If
    
End Sub

Private Sub Finish_Click()
Dim rc
Dim slValue, slData As String
Dim Msg As String
Dim FileHandle
Dim f

    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\LocalWEB\Settings"
    slValue = "HomePage"
    slData = HomePage.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "LogFile"
    slData = LogFile.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "EnableEvents"
    slData = Str(EnableLogging.Value)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "LogFormat"
    slData = LogFormat.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    slValue = "AllowCGI"
    slData = Str(AllowCGI.Value)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    Msg = "The new LocalWEB settings have been saved." + vbCrLf + vbCrLf
    Msg = Msg & "You do not have to reload LocalWEB for the settings to take effect."
    MsgBox Msg, vbOKOnly + vbInformation, "LocalWEB Settings"
    
    Form1.HomePageLoc.Caption = HomePage.Text
    
    Rem ** Check IP Exception List, If items, save to EXCEPT.LST **
    
    If ExceptionList.ListCount > 0 Then
      If FileExists(App.Path & "\except.lst") Then
         Kill (App.Path & "\except.lst")
      End If
      FileHandle = FreeFile
      Open App.Path & "\except.lst" For Output As #FileHandle
      Dim i
      For i = 0 To ExceptionList.ListCount - 1
        Write #FileHandle, ExceptionList.List(i)
      Next i
      Close #FileHandle
    Else
      Kill App.Path & "\except.lst"
    End If
    
    'rewrite MIME type List
    
    FileHandle = FreeFile
    If FileExists(App.Path & "\mime.lst") Then
      Kill App.Path & "\mime.lst"
    End If
    
    Open App.Path & "\mime.lst" For Output As #FileHandle
    For f = 0 To MIMEList.ListCount - 1
      Write #FileHandle, MIMEList.List(f)
    Next f
    Close #FileHandle
    
    
End Sub

Function FileExists(FileName As String) As Boolean
Dim TempAttr As Integer

On Error GoTo ErrorFileExist

    TempAttr = GetAttr(FileName)
    FileExists = ((TempAttr And vbDirectory) = 0)
    GoTo ExitFileExist
    
ErrorFileExist:
    FileExists = False
    Resume ExitFileExist
    
ExitFileExist:
    On Error GoTo 0
    
End Function
Private Sub Form_Load()
Dim FileHandle
Dim Header As String

        
    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\LocalWEB\Settings"
    'Modified by STF from "Not Set" to App.Path
    HomePage.Text = GetRegValue(hKey, SubKey, "HomePage", App.Path)
    'Modified by STF from "Not Set" to "index.htm"
    'IndexPage.Text = GetRegValue(hKey, SubKey, "IndexPage", "index.htm")
    'Modified by STF from "Not Set" to App.Path
    LogFile.Text = GetRegValue(hKey, SubKey, "LogFile", App.Path)
    EnableLogging.Value = Val(GetRegValue(hKey, SubKey, "EnableEvents", "0"))
    AllowCGI.Value = Val(GetRegValue(hKey, SubKey, "AllowCGI", "0"))
    'Modified by STF from "Not Set" to "Common Log File Format"
    LogFormat.Text = GetRegValue(hKey, SubKey, "LogFormat", "Common Log File Format")
    
    Rem ** see if event logging is enabled **
    
    If EnableLogging.Value = 0 Then
      LogFormat.Enabled = False
    Else
      LogFormat.Enabled = True
    End If
    
    Rem ** read Exception IP List **
    
    If FileExists(App.Path & "\except.lst") Then
      
      Dim IPAddress As String
      FileHandle = FreeFile
      Open App.Path & "\except.lst" For Input As #FileHandle
      While Not EOF(FileHandle)
        Input #FileHandle, IPAddress
        ExceptionList.AddItem IPAddress
      Wend
      Close #FileHandle
    End If
        
    'populate combo box with format types
    
    LogFormat.AddItem "Common Log File Format"
    LogFormat.AddItem "NCSA Format"
    LogFormat.AddItem "LocalWEB Format"
    
    'setup MIME List Box
   
    If FileExists(App.Path & "\mime.lst") Then
      FileHandle = FreeFile
      Open App.Path & "\mime.lst" For Input As #FileHandle
      While Not EOF(FileHandle)
        Input #FileHandle, Header
        MIMEList.AddItem Header
      Wend
      Close #FileHandle
    End If
    
    StepCount = 0
    LastStep = 4
    DisplayStep StepCount, LastStep

End Sub

Private Sub Next_Click()

    Frame1(StepCount).Visible = False
    StepCount = StepCount + 1
    DisplayStep StepCount, LastStep
End Sub

Private Sub TreeView1_Click()

    MsgBox (TreeView1.SelectedItem.Text)
End Sub
