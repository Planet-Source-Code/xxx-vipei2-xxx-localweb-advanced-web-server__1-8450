VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form form1 
   Caption         =   "LocalWEB"
   ClientHeight    =   2760
   ClientLeft      =   6360
   ClientTop       =   2775
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet DNS4EVERconnect 
      Left            =   2700
      Top             =   2850
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSWinsockLib.Winsock telnet 
      Left            =   2055
      Top             =   2790
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1455
      Top             =   2790
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer DNS4everupdate 
      Left            =   885
      Top             =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   330
      Top             =   2775
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2730
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4815
      _Version        =   327680
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Server.frx":030A
      Tab(0).ControlCount=   10
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "homepageloc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "pageaddress"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "serverstart"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "connections"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command3"
      Tab(0).Control(9).Enabled=   0   'False
      TabCaption(1)   =   "Local Details"
      TabPicture(1)   =   "Server.frx":0326
      Tab(1).ControlCount=   3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CGISUPPORT"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Image1"
      Tab(1).Control(2).Enabled=   0   'False
      TabCaption(2)   =   "Logging and Monitoring"
      TabPicture(2)   =   "Server.frx":0342
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Totalrequests"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label14"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Telnetsupport"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "statsbox"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "LoadMonitor"
      Tab(2).Control(5).Enabled=   0   'False
      TabCaption(3)   =   "DNS4EVER"
      TabPicture(3)   =   "Server.frx":035E
      Tab(3).ControlCount=   10
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command2"
      Tab(3).Control(0).Enabled=   -1  'True
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "statusreturned"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label11"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "DNS4everenabled"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "dns4everdomainname"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label9"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label8"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label7"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label6"
      Tab(3).Control(9).Enabled=   0   'False
      Begin VB.CommandButton Command3 
         Caption         =   "Options"
         Height          =   375
         Left            =   4410
         TabIndex        =   28
         Top             =   2265
         Width           =   990
      End
      Begin ComctlLib.ProgressBar LoadMonitor 
         Height          =   210
         Left            =   -74895
         TabIndex        =   27
         Top             =   2475
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.ListBox statsbox 
         Height          =   1620
         Left            =   -74895
         TabIndex        =   20
         Top             =   900
         Width           =   5415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start DNS4EVER Support"
         Height          =   285
         Left            =   -73515
         TabIndex        =   19
         Top             =   1980
         Width           =   2580
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Stop"
         Height          =   300
         Left            =   1950
         TabIndex        =   5
         Top             =   2340
         Width           =   1695
      End
      Begin VB.Label CGISUPPORT 
         BackStyle       =   0  'Transparent
         Caption         =   "disabled"
         Height          =   240
         Left            =   -73575
         TabIndex        =   26
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "CGI Support:"
         Height          =   330
         Left            =   -74850
         TabIndex        =   25
         Top             =   630
         Width           =   1290
      End
      Begin VB.Image Image1 
         Height          =   1830
         Left            =   -72240
         Picture         =   "Server.frx":037A
         Top             =   585
         Width           =   2595
      End
      Begin VB.Label Telnetsupport 
         Caption         =   "Enabled"
         Height          =   285
         Left            =   -70620
         TabIndex        =   24
         Top             =   525
         Width           =   930
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Telnet:"
         Height          =   240
         Left            =   -72045
         TabIndex        =   23
         Top             =   540
         Width           =   1350
      End
      Begin VB.Label Totalrequests 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   225
         Left            =   -72540
         TabIndex        =   22
         Top             =   555
         Width           =   720
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Requests on server:"
         Height          =   300
         Left            =   -74820
         TabIndex        =   21
         Top             =   555
         Width           =   2220
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   300
         Left            =   -74775
         TabIndex        =   18
         Top             =   1575
         Width           =   615
      End
      Begin VB.Label statusreturned 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   -74025
         TabIndex        =   17
         Top             =   1575
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "for DNS4ever"
         Height          =   225
         Left            =   -72150
         TabIndex        =   16
         Top             =   795
         Width           =   1800
      End
      Begin VB.Label DNS4everenabled 
         BackStyle       =   0  'Transparent
         Caption         =   "Disabled"
         Height          =   270
         Left            =   -73020
         TabIndex        =   15
         Top             =   795
         Width           =   915
      End
      Begin VB.Label dns4everdomainname 
         BackStyle       =   0  'Transparent
         Caption         =   "dns4everdomainname"
         Height          =   300
         Left            =   -73350
         TabIndex        =   14
         Top             =   1185
         Width           =   2310
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Domain Name:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   13
         Top             =   1185
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Support is currently"
         Height          =   255
         Left            =   -74775
         TabIndex        =   12
         Top             =   795
         Width           =   1890
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "       - DNS4EVER Details -"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73845
         TabIndex        =   11
         Top             =   405
         Width           =   3330
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Visit http://www.dns4ever.com for details on this service!"
         Height          =   270
         Left            =   -74730
         TabIndex        =   10
         Top             =   2400
         Width           =   5070
      End
      Begin VB.Label connections 
         BackStyle       =   0  'Transparent
         Caption         =   "available connections"
         Height          =   255
         Left            =   2235
         TabIndex        =   9
         Top             =   1980
         Width           =   2190
      End
      Begin VB.Label serverstart 
         BackStyle       =   0  'Transparent
         Caption         =   "serverstart"
         Height          =   240
         Left            =   1995
         TabIndex        =   8
         Top             =   1500
         Width           =   1770
      End
      Begin VB.Label pageaddress 
         BackStyle       =   0  'Transparent
         Caption         =   "pageaddress"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2505
         TabIndex        =   7
         Top             =   1110
         Width           =   1860
      End
      Begin VB.Label homepageloc 
         BackStyle       =   0  'Transparent
         Caption         =   "Homepageloc"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1770
         TabIndex        =   6
         Top             =   630
         Width           =   2565
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Connections:"
         Height          =   225
         Left            =   210
         TabIndex        =   4
         Top             =   1980
         Width           =   2385
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Started On:"
         Height          =   270
         Left            =   210
         TabIndex        =   3
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Page Available at:"
         Height          =   345
         Left            =   210
         TabIndex        =   2
         Top             =   1095
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Directory:"
         Height          =   285
         Left            =   195
         TabIndex        =   1
         Top             =   615
         Width           =   1500
      End
   End
   Begin VB.Menu menublah 
      Caption         =   "BigPopUP"
      Visible         =   0   'False
      Begin VB.Menu menuserver 
         Caption         =   "&Server"
         Begin VB.Menu menuexit 
            Caption         =   "E&xit..."
         End
      End
      Begin VB.Menu menuproperties 
         Caption         =   "&Properties"
         Begin VB.Menu mnuprops 
            Caption         =   "&Properties"
         End
      End
      Begin VB.Menu mpopupsys 
         Caption         =   "&SysTray"
         Visible         =   0   'False
         Begin VB.Menu mstartstop 
            Caption         =   "&Stop"
         End
         Begin VB.Menu mpoprestore 
            Caption         =   "&Restore"
         End
         Begin VB.Menu mpopexit 
            Caption         =   "&Exit"
         End
      End
      Begin VB.Menu Menuhelp 
         Caption         =   "&Help"
         Begin VB.Menu menucontents 
            Caption         =   "&Contents.."
         End
         Begin VB.Menu menutip 
            Caption         =   "&Tip of the Day"
         End
         Begin VB.Menu menuabout 
            Caption         =   "&About..."
         End
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LocalWEBRunning As Boolean
Private gSocketInstance
Private ConnectionsUsed
Private AvailableConnections
Private Const MAX_CONNECTIONS = 100
Private LogType As String
Private LogFileHandle
Private LoggingEnabled
Private IndexPage As String

Rem **** Telnet Server Variables ****

Dim Password As String
Dim AcceptedPassword As Boolean
Dim SuccessLogin As Boolean
Dim UserCommand As String
Dim StoredPassword As String
Dim TelnetServerEnabled As String

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

Rem *********************************





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

Private Sub Command1_Click()

If Command1.Caption = "&Stop" Then
       Winsock1(0).Close
       Command1.Caption = "&Start"
       'ButtonMessage.Caption = "Click 'Start' to make the website available"
       form1.Caption = "LocalWEB - Website Unavailable"
       LocalWEBRunning = False
    Else
       Winsock1(0).LocalPort = 80
       Winsock1(0).Protocol = sckTCPProtocol
       Winsock1(0).Listen
       Command1.Caption = "&Stop"
       'ButtonMessage.Caption = "Click 'Stop' to make the website unavailable"
       form1.Caption = "LocalWEB - Website Available"
       LocalWEBRunning = True
    End If
       
'Added by STF
'Added for System Tray
mstartstop.Caption = Command1.Caption
'End STF Add

End Sub

Private Sub Command2_Click()
Dim ReturnCode As String
Dim DomainName As String
Dim UserName As String
Dim Password As String
Dim IPAddress As String
Dim UpdateInterval As String
Dim URL As String

    IPAddress = Winsock1(0).LocalIP
    DomainName = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "DNS4EVERDomainName", "<not set>")
    UserName = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "DNS4EVERUserName", "<not set")
    Password = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "DNS4EVERPassword", "<not set>")
    UpdateInterval = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "DNS4EVERRefreshRate", "60")
    
    If DomainName = "<not set>" Or UserName = "<not set>" Or Password = "<not set>" Then
      MsgBox "DNS4EVER is not configured correctly"
      Exit Sub
    End If
    
    If Command2.Caption = "Start DNS4EVER Support" Then
      Command2.Caption = "Stop DNS4EVER Support"
      URL = "http://www.dns4ever.com/sys/u.cgi?d=" & DomainName & "&u=" & UserName & "&p=" & Password & "&i=" & IPAddress
      ReturnCode = DNS4EVERconnect.OpenURL(URL)
      If Left(ReturnCode, 1) = "0" Then
        statusreturned.Caption = "Invalid/Failed"
      Else
        If Left(ReturnCode, 1) = "1" Then
          statusreturned.Caption = "OK"
        Else
          If Left(ReturnCode, 1) = "2" Then
            statusreturned.Caption = "Trial Expired"
          End If
        End If
      End If
      
      If Val(UpdateInterval) * 1000 > 65535 Then
        UpdateInterval = "60"
      End If
      DNS4everupdate.Interval = Val(UpdateInterval) * 1000
      DNS4everupdate.Enabled = True
    Else
      Command2.Caption = "Start DNS4EVER Support"
      DNS4everupdate.Enabled = False
      statusreturned.Caption = ""
    End If
    
      
End Sub

Private Sub Command3_Click()
PopupMenu menublah
End Sub

Private Sub DNS4EVERUpdate_Timer()
Dim ReturnCode As String
Dim DomainName As String
Dim UserName As String
Dim Password As String
Dim IPAddress As String
Dim URL As String

    IPAddress = Winsock1(0).LocalIP
    DomainName = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "DNS4EVERDomainName", "<not set>")
    UserName = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "DNS4EVERUserName", "<not set")
    Password = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "DNS4EVERPassword", "<not set>")
    
    URL = "http://www.dns4ever.com/sys/u.cgi?d=" & DomainName & "&u=" & UserName & "&p=" & Password & "&i=" & IPAddress
    ReturnCode = DNS4EVERconnect.OpenURL(URL)
    statusreturned.Caption = Left(ReturnCode, 1)
   
    

End Sub

Private Sub Form_Load()
Dim Response
Dim Msg As String
Dim SCounter As Integer
Dim TipsEnabled As String
Dim LogPath As String
    
    
    'Bit of Fault Tolerance, if default.lst or mime.lst is missing
    'make LocalWEB build a simple set of files
    
    If Not FileExists(App.Path & "\mime.lst") Then
      BuildNewMimeList
    End If
    
    If Not FileExists(App.Path & "\default.lst") Then
      BuildNewDefaultList
    End If
    
    
    
    'Is there anything in the registry, if no run the setup Wizard
    
    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\LocalWEB\Settings"
    If Not CheckRegistry Then
      Msg = "LocalWeb has not been configured." + vbCrLf + vbCrLf
      Msg = Msg & "LocalWEB will configure itself with default values." + vbCrLf + vbCrLf
      Msg = Msg & "Configure LocalWEB now?"
      Response = MsgBox(Msg, vbQuestion + vbYesNo, "LocalWEB")
      If Response = vbYes Then
        CreateRegKeys
      Else
        Stop
        Exit Sub
      End If
    End If
      
    'show the tip of the day if it is enabled
    
    'Modified by STF from "Not Set" to 1. This forces it to ON if not set
    TipsEnabled = GetRegValue(hKey, SubKey, "TipOfTheDay", 1)
   
    
   ' If TipsEnabled Then
   '   mnuTip.Checked = True
   '   frmTip.Show 1
   ' Else
   '   mnuTip.Checked = False
   ' End If
    
    Rem ****************** Telnet Server Listener *****************
    
    telnet.LocalPort = 23
    telnet.Listen
    Password = ""
    UserCommand = ""
    SuccessLogin = False
    AcceptedPassword = False
    
   
    
    
    Rem ****************** HTTP Server Listener *******************
    
    ConnectionsUsed = 0
    LoadMonitor.Min = 0
    LoadMonitor.Value = Min
    Winsock1(0).LocalPort = 80
    Winsock1(0).Protocol = sckTCPProtocol
    Winsock1(0).Listen
    LocalWEBRunning = True
    
    form1.Caption = "LocalWEB - Website Available"
    
    For SCounter = 1 To MAX_CONNECTIONS
      Load Winsock1(SCounter)
    Next SCounter
    
    
    'Open LogFile, if event logging is enabled
    
    LoggingEnabled = GetRegValue(hKey, SubKey, "EnableEvents", "1")
    If Val(LoggingEnabled) = 1 Then
      LogPath = GetRegValue(hKey, SubKey, "LogFile", "c:\")
      LogType = GetRegValue(hKey, SubKey, "LogFormat", "LocalWEB Format")
      LogFileHandle = FreeFile
      Open LogPath & "\localweb.log" For Output As #LogFileHandle
    End If
   
    AvailableConnections = MAX_CONNECTIONS
    connections.Caption = Str(AvailableConnections)
    
    Totalrequests.Caption = "0"
    serverstart.Caption = Format(Date, "dddd, mmmm d yyyy") & "  at " & Format(Time, "hh:mm:ss AMPM")
   
    homepageloc.Caption = GetRegValue(hKey, SubKey, "HomePage", "Not Set")
    IndexPage = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "IndexPage", "index.htm")
    pageaddress.Caption = "http://" & Winsock1(0).LocalIP
    
    If Val(GetRegValue(hKey, SubKey, "AllowCGI", "0")) = 0 Then
      CGISUPPORT.Caption = "Disabled"
    Else
      CGISUPPORT.Caption = "Enabled"
    End If
    If Val(GetRegValue(hKey, SubKey, "TelnetEnabled", "0")) = 0 Then
      Telnetsupport.Caption = "Disabled"
    Else
      Telnetsupport.Caption = "Enabled"
    End If
    
    'load dns4ever information
    
    If Val(GetRegValue(hKey, SubKey, "DNS4EVEREnabled", "0")) = 0 Then
      DNS4everenabled.Caption = "Disabled"
      Command2.Enabled = False
    Else
      DNS4everenabled.Caption = "Enabled"
      Command2.Enabled = True
    End If
    If GetRegValue(hKey, SubKey, "DNS4EVERDomainName", "<not set>") = "<not set>" Then
      dns4everdomainname = "<not set>"
    Else
      dns4everdomainname.Caption = "http://" & GetRegValue(hKey, SubKey, "DNS4EVERDomainName", "<not set>")
    End If
    'disable the update timer
    DNS4everupdate.Enabled = False
    
    statsbox.AddItem "LocalWEB Server IP Address " & Winsock1(0).LocalIP & " listening on Port " & Winsock1(0).LocalPort
    statsbox.AddItem "LocalWEB Ready......."
    
'Added by STF
'For System Tray
       'the form must be fully visible before calling Shell_NotifyIcon
       Me.Show
       Me.Refresh
       With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "LocalWEB: " & Winsock1(0).LocalIP & vbNullChar
       End With
       Shell_NotifyIcon NIM_ADD, nid
      
'End STF Add
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Added by STF
'For System Tray
      'this procedure receives the callbacks from the System Tray icon.
      Dim Result As Long
      Dim Msg As Long
       'the value of X will vary depending upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        Msg = X
       Else
        Msg = X / Screen.TwipsPerPixelX
       End If
       Select Case Msg
        Case WM_LBUTTONUP        '514 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mpopupsys
         'Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mpopupsys
         'Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mpopupsys
       End Select
'End STF Add
End Sub


Private Sub Form_Resize()
'Added by STF
'Added for System Tray
       'this is necessary to assure that the minimized window is hidden
       If Me.WindowState = vbMinimized Then Me.Hide
'End STF Add
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim SCount As Integer

'Added by STF
'For System Tray
       'this removes the icon from the system tray
       Shell_NotifyIcon NIM_DELETE, nid
'End STF Add

    For SCount = 1 To MAX_CONNECTIONS
      Unload Winsock1(SCount)
    Next
    
    
End Sub


Private Sub mnuDocuments_Click()

    frmDocuments.Show 1
    
End Sub

Private Sub mnuErrorPages_Click()

    frmServerErrors.Show 1
    
End Sub

Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub mnuProps_Click()

    frmServerProperties.Show 1
    
End Sub

Private Sub mnuTelnet_Click()

    frmTelnet.Show 1
End Sub

Private Sub mPopExit_Click()
'Added by STF
'Added for System Tray
       'called when user clicks the popup menu Exit command
Call MenuExit_Click
'End STF Add
End Sub

Private Sub mPopRestore_Click()
'Added by STF
'Added for System Tray
       'called when the user clicks the popup menu Restore command
       Me.WindowState = vbNormal
       Result = SetForegroundWindow(Me.hwnd)
       Me.Show
'End STF Add
End Sub



Private Sub HomePageLoc_Click()
Shell "explorer /e," & Trim(homepageloc.Caption), 1

End Sub

Private Sub MenuAbout_Click()

    frmAbout.Show 1
End Sub

Private Sub MenuExit_Click()
Dim rc
Dim slData, slValue


    'set TipOfTheDay flag in registry
    
    SubKey = "SOFTWARE\LocalWEB\Settings"
    slValue = "TipOfTheDay"
    'If mnuTip.Checked = True Then
    '  slData = "1"
    'Else
    '  slData = "0"
    'End If
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
        
    If Val(LoggingEnabled) = 1 Then
      Close #LogFileHandle
    End If
    Unload Me
End Sub

Private Sub MIMETypes_Click()

    frmMIME.Show 1
End Sub

Private Sub mnuMIME_Click()

    frmMIME.Show 1
End Sub

Private Sub mnuTip_Click()

    If mnuTip.Checked = True Then
      mnuTip.Checked = False
    Else
      mnuTip.Checked = True
    End If
    
End Sub

Private Sub mnuWizard_Click()

    frmWizard.Show 1
    
End Sub


Private Sub mStartStop_Click()
Call Command1_Click
End Sub

Private Sub PageAddress_Click()
    ShowURL Trim(pageaddress.Caption)
End Sub

Private Sub Telnet_Close()

    telnet.Close
    telnet.LocalPort = 23
    telnet.Listen
    Password = ""
    UserCommand = ""
    SuccessLogin = False
    AcceptedPassword = False
    
End Sub

Private Sub Telnet_ConnectionRequest(ByVal requestID As Long)

    If telnet.State <> sckClosed Then
      telnet.Close
    End If
    
    telnet.Accept requestID
    
    TelnetServerEnabled = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "TelnetEnabled", "0")
    
    If Val(TelnetServerEnabled) = 0 Then
      telnet.SendData "LocalWEB HTTP Server" & vbCrLf & vbCrLf
      telnet.SendData "LocalWEB Telnet Server is DISABLED"
    Else
      telnet.SendData "LocalWEB HTTP Server" & vbCrLf & vbCrLf
      telnet.SendData "Password:"
    End If
    
    
End Sub

Private Sub Telnet_DataArrival(ByVal bytesTotal As Long)
Dim str1 As String
    
    telnet.GetData str1
    If SuccessLogin Then
      If Asc(str1) = 13 Then
        If UCase(UserCommand) Like "HELP" Then
          telnet.SendData vbCrLf & vbCrLf & "Commands available in Remote Admin:" & vbCrLf & vbCrLf
          telnet.SendData "help     -- display commands available in Remote Admin." & vbCrLf
          telnet.SendData "exit     -- logout of Remote Admin." & vbCrLf
          telnet.SendData "stat     -- show if LocalWEB is stopped or started." & vbCrLf
          telnet.SendData "enable   -- enable LocalWEB." & vbCrLf
          telnet.SendData "disable  -- disable LocalWEB." & vbCrLf
          telnet.SendData "site     -- show site locations and server address." & vbCrLf
          telnet.SendData "counters -- show LocalWEB statistics counters." & vbCrLf
          
          telnet.SendData vbCrLf & ">"
          UserCommand = ""
        ElseIf UCase(UserCommand) Like "EXIT" Then
          telnet.SendData vbCrLf & "Exit request from node " & telnet.RemoteHostIP
          Telnet_Close
          
        ElseIf UCase(UserCommand) Like "STAT" Then
          If LocalWEBRunning Then
            telnet.SendData vbCrLf & "LocalWEB is currently ENABLED." & vbCrLf
          Else
            telnet.SendData vbCrLf & "LocalWEB is Currently DISABLED." & vbCrLf
          End If
          UserCommand = ""
          telnet.SendData vbCrLf & ">"
        ElseIf UCase(UserCommand) = "ENABLE" Then
          If LocalWEBRunning Then
            telnet.SendData vbCrLf & "LocalWEB is already ENABLED." & vbCrLf
          Else
            Winsock1(0).LocalPort = 80
            Winsock1(0).Protocol = sckTCPProtocol
            Winsock1(0).Listen
            Command1.Caption = "&Stop"
            mstartstop.Caption = "&Stop"
            ButtonMessage.Caption = "Click 'Stop' to make the website unavailable"
            form1.Caption = "LocalWEB - Website Available"
            LocalWEBRunning = True
            telnet.SendData vbCrLf & "LocalWEB is now ENABLED."
          End If
          UserCommand = ""
          telnet.SendData vbCrLf & ">"
        ElseIf UCase(UserCommand) = "DISABLE" Then
          If LocalWEBRunning = False Then
            telnet.SendData vbCrLf & "LocalWEB is already DISABLED." & vbCrLf
          Else
            Winsock1(0).Close
            Command1.Caption = "&Start"
            mstartstop.Caption = "&Start"
            ButtonMessage.Caption = "Click 'Start' to make the website available"
            form1.Caption = "LocalWEB - Website Unavailable"
            LocalWEBRunning = False
            telnet.SendData vbCrLf & "LocalWEB is now DISABLED." & vbCrLf
          End If
          UserCommand = ""
          telnet.SendData vbCrLf & ">"
        ElseIf UCase(UserCommand) = "SITE" Then
          telnet.SendData vbCrLf & vbCrLf & "Index page location: " & form1.homepageloc & vbCrLf
          telnet.SendData "Server IP Address: " & form1.pageaddress & vbCrLf
          UserCommand = ""
          telnet.SendData vbCrLf & ">"
        ElseIf UCase(UserCommand) = "COUNTERS" Then
          telnet.SendData vbCrLf & vbCrLf & "Connections available: " & form1.connections & vbCrLf
          telnet.SendData "Total requests so far: " & form1.Totalrequests & vbCrLf
          telnet.SendData "LocalWEB started on: " & form1.serverstart & vbCrLf
          telnet.SendData vbCrLf & ">"
          UserCommand = ""
        Else
          telnet.SendData vbCrLf & vbCrLf & "The command is not recognised." & vbCrLf
          UserCommand = ""
          telnet.SendData vbCrLf & ">"
        End If
      Else
        telnet.SendData str1
        UserCommand = UserCommand & str1
      End If
    Else
      'get user authentication
      StoredPassword = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "TelnetPassword", "Password")
      If Not AcceptedPassword And Asc(str1) = 13 Then
        If UCase(Password) = UCase(StoredPassword) Then
          SuccessLogin = True
          AcceptedPassword = True
          telnet.SendData vbCrLf & "LocalWEB Remote Admin version 1.0.0"
          telnet.SendData vbCrLf & "Login accepted from node " & telnet.RemoteHostIP & vbCrLf
          telnet.SendData "To list the commands available type help at the prompt." & vbCrLf
          telnet.SendData vbCrLf & ">"
        Else
          Password = ""
          telnet.SendData vbCrLf & "Invalid password from node " & telnet.RemoteHostIP & vbCrLf & vbCrLf
          telnet.SendData "LocalWEB HTTP Server" & vbCrLf & vbCrLf
          telnet.SendData "Password:"
          Exit Sub
        End If
      Else
        Password = Password & str1
      End If
    End If
    
    
          
          
End Sub

Private Sub Timer1_Timer()
Dim X As Integer

    For X = 1 To Winsock1.Count - 1
      If Winsock1(X).State <> sckClosed And Val(Winsock1(X).Tag) < Timer And Val(Winsock1(X).Tag) > 0 Then
        Winsock1(X).Close
      End If
    Next X
End Sub

Private Sub Winsock1_Close(Index As Integer)

    Winsock1(Index).Close
   
    
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim TotalReq As Integer
Dim X As Integer
Dim DenyList As String
Dim FreeHandle
Dim Blocked As Boolean

    TotalReq = Val(Totalrequests.Caption)
    TotalReq = TotalReq + 1
    Totalrequests.Caption = Str(TotalReq)
    
    For X = 1 To Winsock1.Count - 1
      If X > Winsock1.Count - 1 Then Exit For
      If Winsock1(X).State = sckClosed Then
        Winsock1(X).Tag = Timer + 5
        Winsock1(X).Accept requestID
        'check if IP address is a blocked one
        If FileExists(App.Path & "\deny.lst") Then
          FreeHandle = FreeFile
          Open App.Path & "\deny.lst" For Input As #FreeHandle
          Blocked = False
          Do While Not EOF(FreeHandle)
            Input #FreeHandle, DenyList
            If Winsock1(Index).RemoteHostIP = DenyList Then
              Blocked = True
              Exit Do
            End If
          Loop
          Close #FreeHandle
          If Blocked = True Then
            Winsock1(X).SendData "HTTP/1.0 200 OK"
            Winsock1(X).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
            Winsock1(X).SendData "<HTML>"
            Winsock1(X).SendData "<BODY>"
            Winsock1(X).SendData "<TITLE>404.4 Access Denied</TITLE>"
            Winsock1(X).SendData "<h1>404.4 Not Found</h1>"
            Winsock1(X).SendData "<p>Access to this site has been blocked for this machine.</p>" + vbCrLf + vbCrLf
            Winsock1(X).SendData "<p><em>LocalWEB HTTP Server</em></p>"
            Winsock1(X).SendData "</BODY>"
            Winsock1(X).SendData "</HTML>"
          End If
        End If
        ConnectionsUsed = ConnectionsUsed + 1
        LoadMonitor.Value = (ConnectionsUsed / MAX_CONNECTIONS) * 100
        Exit For
      End If
    Next X
           
    statsbox.AddItem "Request From " & Winsock1(Index).RemoteHostIP
    
    
      
        
          
    
    
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim FileName As String
Dim clstring As String
Dim clstringpass1 As String
Dim HTMFile As String
Dim MyChar, spointer
Dim TextLine As String
Dim MIME As String
Dim WebROOT As String
Dim szBuffer As String
Dim FileExt As String
Dim FileHandle
Dim HTMLHandle
Dim DListHandle
Dim ErrorHandle, Error404File As String
Dim MIMEExt, MIMEHeader, MIMEFile As String
Dim HeaderOK As Boolean
Dim StatusCode As String
Dim BytesSent As Long

Dim HTML As String
Dim Header As String
Dim FooterCode As String
'virtual directory variables
Dim Pagename As String
Dim dirname As String
Dim length As Integer
Dim dot As Integer
Dim VirtualHandle
Dim VirtualDir As String
Dim VirtualDirs As String
Dim RealDir As String
Dim VirtualFound As Boolean
Dim PathAndScript As String

    WebROOT = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "HomePage", "c:\www\root\")
    WebROOT = WebROOT + "\"
    FooterCode = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "FooterCode", "<p><font color=&quot;#0000FF&quot; size=&quot;1&quot;>This site is served by LocalWEB</font></p>")
    
    
    clstring = ""
    clstringpass1 = ""
    MyChar = ""
    HTMFile = ""
    


    Winsock1(Index).GetData clstring, vbString
    DoEvents
    
    
    Rem ** Parse GET command to get file name **
    Rem ** bit we want is before the HTTP bit, so get start position of HTTP **
    
 If Left(clstring, 3) = "GET" Then
    Winsock1(Index).Tag = Timer + 5
    spointer = InStr(1, clstring, "HTTP")
    clstringpass1 = Left(clstring, spointer - 2)
    spointer = InStr(1, clstringpass1, "/")
    For f = spointer + 1 To Len(clstringpass1)
      HTMFile = HTMFile & Mid$(clstringpass1, f, 1)
    Next
    
    
    'check if there is a page name supplied
    'if not then check for the defaults
    'and we must check the virtual directories
    
    If Len(HTMFile) = 0 Then
      DListHandle = FreeFile
      Open App.Path & "\default.lst" For Input As #DListHandle
      Do Until EOF(DListHandle)
        Input #DListHandle, HTMFile
        If FileExists(WebROOT + HTMFile) Then
          Exit Do
        End If
      Loop
      Close #DListHandle
    Else
    'VIRTUAL DIRECTORY CHECK
      'convert any '\' to '/' for correct parsing
      For f = 1 To Len(HTMFile)
        If Mid(HTMFile, f, 1) = "\" Then
          Mid(HTMFile, f, 1) = "/"
        End If
      Next
      f = InStr(1, HTMFile, ".")
      If f > 0 Then
        'we have a pagename
        For i = f To 1 Step -1
          If Mid$(HTMFile, i, 1) = "/" Then
            Exit For
          End If
        Next i
        For f1 = i + 1 To Len(HTMFile)
          Pagename = Pagename & Mid$(HTMFile, f1, 1)
        Next f1
        For f1 = 1 To i
          dirname = dirname & Mid$(HTMFile, f1, 1)
        Next f1
      Else
        dirname = HTMFile
      End If
    
      'now convert virtual dirs and do a pagename check
      If Len(dirname) > 0 Then
         If Right(dirname, 1) <> "/" Then
           dirname = dirname & "/"
         End If
         VirtualFound = False
         VirtualHandle = FreeFile
         Open App.Path & "\vdirs.lst" For Input As #VirtualHandle
         Do While Not EOF(VirtualHandle)
           VirtualDir = ""
           RealDir = ""
           Input #VirtualHandle, VirtualDirs
           spointer = InStr(1, VirtualDirs, ",")
           VirtualDir = Left(VirtualDirs, spointer - 1)
           For f = spointer + 1 To Len(VirtualDirs)
             RealDir = RealDir & Mid$(VirtualDirs, f, 1)
           Next f
           If VirtualDir = dirname Then
            VirtualFound = True
              Exit Do
           End If
         Loop
         Close #VirtualHandle
         If VirtualFound Then
          WebROOT = RealDir & "\"
         Else
         WebROOT = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "HomePage", "c:\www\root\")
         WebROOT = WebROOT + "\"
         End If
      End If
      If Len(Pagename) = 0 Then
        DListHandle = FreeFile
        Open App.Path & "\default.lst" For Input As #DListHandle
        Do Until EOF(DListHandle)
          Input #DListHandle, HTMFile
          If FileExists(WebROOT + HTMFile) Then
            Exit Do
          End If
        Loop
        Close #DListHandle
      Else
        HTMFile = Pagename
      End If
    End If
    
    Rem *** Check to see if its a call to a CGI Script ***
    
    spointer = InStr(1, HTMFile, "cgi-bin")
    If spointer = 1 Then        'Call to CGI
      'cgi directory is always underneath Website root, so hardcode it
      WebROOT = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "HomePage", "c:\www\root\")
      HTMFile = Left(HTMFile, Len(HTMFile) - 1)
      statsbox.AddItem "Executing CGI Script " & HTMFile
      ExecuteCGI HTMFile, WebROOT, Index
    Else
       Rem *** check if HTML File Exists ***
    
       If Not FileExists(WebROOT + HTMFile) Then
         If GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "Error404", "Default") = "Default" Then
           Winsock1(Index).SendData "HTTP/1.0 200 OK"
           Winsock1(Index).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
           Winsock1(Index).SendData "<HTML>"
           Winsock1(Index).SendData "<BODY>"
           Winsock1(Index).SendData "<TITLE>404 Not Found</TITLE>"
           Winsock1(Index).SendData "<h1>404 Not Found</h1>"
           Winsock1(Index).SendData "<p> The requested page could not be found on this server.</p>" + vbCrLf + vbCrLf
           Winsock1(Index).SendData "<p><em>LocalWEB HTTP Server</em></p>"
           Winsock1(Index).SendData "</BODY>"
           Winsock1(Index).SendData "</HTML>"
           StatusCode = "404"
           BytesSent = 0
         Else
           StatusCode = "404"
           BytesSent = 0
           Error404File = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "Error404", "Default")
           ErrorHandle = FreeFile
           Open Error404File For Binary Shared As #ErrorHandle
           HTML = Space(LOF(ErrorHandle))
           Get #ErrorHandle, 1, HTML
           Close #ErrorHandle
           Header = "HTTP/1.0 200 OK" + Chr(13) + Chr(10)
           Header = Header + "Server: LocalWEB" + Chr(13) + Chr(10)
           Header = Header + "Content-Type: " + MIMEHeader + Chr(13) + Chr(10)
           Header = Header + "Accept-Ranges: bytes" + Chr(13) + Chr(10)
           Header = Header + "Content-Length: " + LTrim(Str(Len(HTML))) + Chr(13) + Chr(10)
           Header = Header + Chr(13) + Chr(10)
           Buf = Header + HTML
           Winsock1(Index).SendData Buf
         End If
       Else
         'Winsock1(Index).SendData "HTTP/1.0 200 OK"
         StatusCode = "200"
         BytesSent = FileLen(WebROOT & HTMFile)
         
         
         Rem *** check the MIME.LST for file/header association **
         
         spointer = InStr(1, HTMFile, ".")
         For f = spointer + 1 To Len(HTMFile)
           FileExt = FileExt & Mid$(HTMFile, f, 1)
         Next f
         
         FileHandle = FreeFile
         HeaderOK = False
         
         Open App.Path & "\mime.lst" For Input As #FileHandle
         Do While Not EOF(FileHandle)
           MIMEHeader = ""
           Input #FileHandle, MIMEFile
           spointer = InStr(1, MIMEFile, ",")
           MIMEExt = Left(MIMEFile, spointer - 1)
           For f = spointer + 1 To Len(MIMEFile)
             MIMEHeader = MIMEHeader & Mid$(MIMEFile, f, 1)
           Next f
           If MIMEExt = FileExt Then
             HeaderOK = True
             Exit Do
           End If
         Loop
         Close #FileHandle
         
         If HeaderOK Then
           statsbox.AddItem "Requesting File " & HTMFile
           HTMLHandle = FreeFile
           FileName = WebROOT & HTMFile
           Open FileName For Binary Shared As #HTMLHandle
           HTML = Space(LOF(HTMLHandle))
           Get #HTMLHandle, 1, HTML
           Close #HTMLHandle
           Winsock1(Index).Tag = Timer + 5
           Header = "HTTP/1.0 200 OK" + Chr(13) + Chr(10)
           Header = Header + "Server: LocalWEB" + Chr(13) + Chr(10)
           Header = Header + "Content-Type: " + MIMEHeader + Chr(13) + Chr(10)
           Header = Header + "Accept-Ranges: bytes" + Chr(13) + Chr(10)
           Header = Header + "Content-Length: " + LTrim(Str(Len(HTML))) + Chr(13) + Chr(10)
           Header = Header + Chr(13) + Chr(10)
           Buf = Header + HTML
           
           'Do we need to append footer code?
           
           If Val(GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\LocalWEB\Settings", "FooterEnabled", "1")) = 1 Then
             Buf = Header + HTML + FooterCode
           End If
           
           'Now Serve the Page
           Winsock1(Index).SendData Buf
         End If
       End If
    End If
    
    'write log file
    
    If Val(LoggingEnabled) = 1 Then
      If LogType = "LocalWEB Format" Then
        Print #LogFileHandle, Winsock1(0).RemoteHostIP + "," + Format(Date, "dd/mm/yyyy") + "," + Format(Time, "hh:mm:ss") + ",GET," + StatusCode + "," + Str(BytesSent) + "," + Winsock1(Index).LocalIP + "," + clstringpass1
      Else
        If LogType = "Common Log File Format" Then
          Print #LogFileHandle, Winsock1(0).RemoteHostIP + "," + "-," + "[" + Format(Date, "dd/mm/yyyy") + ":" + Format(Time, "hh:mm:ss") + "]," + clstringpass1 + "," + StatusCode + "," + Str(BytesSent)
        Else
          If LogType = "NCSA Format" Then
            Print #LogFileHandle, Winsock1(0).RemoteHostIP + ",-,-" + "[" + Format(Date, "dd/mm/yyyy") + ":" + Format(Time, "hh:mm:ss") + "]," + clstringpass1 + "," + StatusCode + "," + Str(BytesSent)
          End If
        End If
      End If
    End If
        
End If 'GET check

'End If
    
    
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)

    Winsock1(Index).Close
    Winsock1(Index).Tag = 0
    ConnectionsUsed = ConnectionsUsed - 1
    LoadMonitor.Value = (ConnectionsUsed / MAX_CONNECTIONS) * 100
   
End Sub

