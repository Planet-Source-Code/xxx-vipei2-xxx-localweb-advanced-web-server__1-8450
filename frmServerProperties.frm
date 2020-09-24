VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmServerProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LocalWEB Properties"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmServerProperties.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   11456
      _Version        =   327680
      Tabs            =   10
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Site Location"
      TabPicture(0)   =   "frmServerProperties.frx":0442
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Logging"
      TabPicture(1)   =   "frmServerProperties.frx":045E
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(2)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "CGI Security"
      TabPicture(2)   =   "frmServerProperties.frx":047A
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(3)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "MIME Types"
      TabPicture(3)   =   "frmServerProperties.frx":0496
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(4)"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "Telnet Server"
      TabPicture(4)   =   "frmServerProperties.frx":04B2
      Tab(4).ControlCount=   3
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TNetText"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "TNetImage"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "TNetFrame"
      Tab(4).Control(2).Enabled=   0   'False
      TabCaption(5)   =   "Documents"
      TabPicture(5)   =   "frmServerProperties.frx":04CE
      Tab(5).ControlCount=   4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "DocFrane"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "FooterFrame"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "DocImage"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "DocDescription"
      Tab(5).Control(3).Enabled=   0   'False
      TabCaption(6)   =   "Errors"
      TabPicture(6)   =   "frmServerProperties.frx":04EA
      Tab(6).ControlCount=   5
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command11"
      Tab(6).Control(0).Enabled=   -1  'True
      Tab(6).Control(1)=   "Command2"
      Tab(6).Control(1).Enabled=   -1  'True
      Tab(6).Control(2)=   "ErrorList"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "ErrDesc"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Image9"
      Tab(6).Control(4).Enabled=   0   'False
      TabCaption(7)   =   "Security"
      TabPicture(7)   =   "frmServerProperties.frx":0506
      Tab(7).ControlCount=   5
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Command13"
      Tab(7).Control(0).Enabled=   -1  'True
      Tab(7).Control(1)=   "Command12"
      Tab(7).Control(1).Enabled=   -1  'True
      Tab(7).Control(2)=   "ServerDeny"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Label19"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Image10"
      Tab(7).Control(4).Enabled=   0   'False
      TabCaption(8)   =   "DNS4EVER "
      TabPicture(8)   =   "frmServerProperties.frx":0522
      Tab(8).ControlCount=   1
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame2"
      Tab(8).Control(0).Enabled=   0   'False
      TabCaption(9)   =   "Directories"
      TabPicture(9)   =   "frmServerProperties.frx":053E
      Tab(9).ControlCount=   3
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame9"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Label17"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "Image8"
      Tab(9).Control(2).Enabled=   0   'False
      Begin VB.CommandButton Command13 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -70800
         TabIndex        =   76
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -70800
         TabIndex        =   75
         Top             =   3000
         Width           =   1215
      End
      Begin ComctlLib.ListView ServerDeny 
         Height          =   2175
         Left            =   -74160
         TabIndex        =   74
         Top             =   2520
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3836
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   -71640
         TabIndex        =   70
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Default"
         Height          =   375
         Left            =   -70440
         TabIndex        =   69
         Top             =   5640
         Width           =   1095
      End
      Begin ComctlLib.ListView ErrorList 
         Height          =   3375
         Left            =   -74640
         TabIndex        =   68
         Top             =   2040
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Error Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "HTML Page Location"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame TNetFrame 
         Caption         =   " Settings "
         Height          =   2055
         Left            =   -74760
         TabIndex        =   62
         Top             =   2400
         Width           =   5655
         Begin VB.CheckBox TelnetEnabled 
            Caption         =   "Enable LocalWEB's Telnet Server"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   360
            Width           =   2775
         End
         Begin VB.Frame TNetFrame2 
            Caption         =   " Telnet Server Password "
            Height          =   855
            Left            =   240
            TabIndex        =   63
            Top             =   840
            Width           =   5175
            Begin VB.TextBox TelnetPassword 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   240
               PasswordChar    =   "*"
               TabIndex        =   64
               Top             =   360
               Width           =   3975
            End
         End
      End
      Begin VB.Frame DocFrane 
         Caption         =   "Default Document List"
         Height          =   2055
         Left            =   -74520
         TabIndex        =   55
         Top             =   1920
         Width           =   5175
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
            TabIndex        =   60
            ToolTipText     =   "Move up list"
            Top             =   720
            Width           =   375
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
            TabIndex        =   59
            ToolTipText     =   "Move down list"
            Top             =   1320
            Width           =   375
         End
         Begin VB.ListBox DefaultList 
            Height          =   1425
            Left            =   600
            TabIndex        =   58
            Top             =   480
            Width           =   2655
         End
         Begin VB.CommandButton AddDocument 
            Caption         =   "&Add..."
            Height          =   375
            Left            =   3480
            TabIndex        =   57
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton RemoveDocument 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   3480
            TabIndex        =   56
            Top             =   1200
            Width           =   1335
         End
      End
      Begin VB.Frame FooterFrame 
         Caption         =   "Page Footer"
         Height          =   1335
         Left            =   -74520
         TabIndex        =   51
         Top             =   4200
         Width           =   5175
         Begin VB.CheckBox FooterEnabled 
            Caption         =   "Enable Page Footer"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.TextBox FooterCode 
            Height          =   285
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   4935
         End
         Begin VB.Label FooterHTML 
            Caption         =   "Footer HTML"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Configured Virtual Directories"
         Height          =   4215
         Left            =   -74760
         TabIndex        =   45
         Top             =   1920
         Width           =   5655
         Begin ComctlLib.ListView VirtualDirectories 
            Height          =   3255
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   5741
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Virtual Directory"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Location"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton Command10 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   3360
            TabIndex        =   48
            Top             =   3720
            Width           =   975
         End
         Begin VB.CommandButton Command9 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   4440
            TabIndex        =   47
            Top             =   3720
            Width           =   975
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Add"
            Height          =   375
            Left            =   2280
            TabIndex        =   46
            Top             =   3720
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5295
         Left            =   -74760
         TabIndex        =   36
         Top             =   840
         Width           =   5655
         Begin VB.Frame Frame8 
            Caption         =   "Authentication Details"
            Height          =   1815
            Left            =   240
            TabIndex        =   40
            Top             =   2520
            Width           =   5175
            Begin VB.TextBox DNS4EVERRefreshRate 
               Height          =   285
               Left            =   2160
               TabIndex        =   72
               Top             =   1440
               Width           =   855
            End
            Begin VB.TextBox DNS4EVERPassword 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   2160
               PasswordChar    =   "*"
               TabIndex        =   3
               Top             =   1080
               Width           =   2535
            End
            Begin VB.TextBox DNS4EVERUserName 
               Height          =   285
               Left            =   2160
               TabIndex        =   2
               Top             =   720
               Width           =   2535
            End
            Begin VB.TextBox DNS4EVERDomainName 
               Height          =   285
               Left            =   2160
               TabIndex        =   1
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label18 
               Caption         =   "Refresh Rate (Seconds):"
               Height          =   255
               Left            =   240
               TabIndex        =   71
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label Label15 
               Caption         =   "Password:"
               Height          =   255
               Left            =   240
               TabIndex        =   43
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label12 
               Caption         =   "Username:"
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label11 
               Caption         =   "Domain Name:"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.CheckBox DNS4EVEREnabled 
            Caption         =   "Enable DNS4EVER Support"
            Height          =   255
            Left            =   240
            TabIndex        =   0
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label Label10 
            Caption         =   "http://www.dns4ever.com"
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
            Height          =   255
            Left            =   2520
            TabIndex        =   39
            Top             =   4800
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "To sign up for this service visit"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   4800
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   $"frmServerProperties.frx":055A
            Height          =   1455
            Left            =   720
            TabIndex        =   37
            Top             =   360
            Width           =   4815
         End
         Begin VB.Image Image6 
            Height          =   480
            Left            =   120
            Picture         =   "frmServerProperties.frx":0724
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4335
         Index           =   4
         Left            =   -74760
         TabIndex        =   31
         Top             =   840
         Width           =   5535
         Begin VB.ListBox MimeList 
            Height          =   2400
            Left            =   240
            TabIndex        =   34
            Top             =   1320
            Width           =   3255
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Add Type..."
            Height          =   375
            Left            =   3720
            TabIndex        =   33
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton Command8 
            Caption         =   "&Remove Type"
            Height          =   375
            Left            =   3720
            TabIndex        =   32
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   240
            Picture         =   "frmServerProperties.frx":0A2E
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label9 
            Caption         =   $"frmServerProperties.frx":0E70
            Height          =   855
            Left            =   840
            TabIndex        =   35
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5175
         Index           =   3
         Left            =   -74760
         TabIndex        =   22
         Top             =   840
         Width           =   5655
         Begin VB.Frame Frame4 
            Caption         =   "Except From The Following IP Addresses"
            Height          =   2535
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   2400
            Width           =   5415
            Begin ComctlLib.ListView ExceptionList 
               Height          =   1335
               Left            =   240
               TabIndex        =   49
               Top             =   960
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   2355
               View            =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   327682
               Icons           =   "ImageList1"
               SmallIcons      =   "ImageList1"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin VB.CommandButton Command6 
               Caption         =   "&Add..."
               Height          =   255
               Left            =   3840
               TabIndex        =   28
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CommandButton Command7 
               Caption         =   "&Remove"
               Height          =   255
               Left            =   3840
               TabIndex        =   27
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Image Image7 
               Height          =   480
               Left            =   120
               Picture         =   "frmServerProperties.frx":0F27
               Top             =   360
               Width           =   480
            End
            Begin VB.Label Label16 
               Caption         =   "The following machines with these IP Addresses will not be able to execute CGI Scripts."
               Height          =   495
               Left            =   840
               TabIndex        =   29
               Top             =   360
               Width           =   4215
            End
         End
         Begin VB.Frame Frame7 
            Height          =   1215
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Width           =   5415
            Begin VB.CheckBox AllowCGI 
               Caption         =   "Allow CGI Scripts To Be Run"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label Label8 
               Caption         =   "If you want LocalWEB to stop execution of all CGI scripts, simply un-check the box above."
               Height          =   495
               Left            =   120
               TabIndex        =   25
               Top             =   600
               Width           =   4935
            End
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   240
            Picture         =   "frmServerProperties.frx":1369
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label7 
            Caption         =   $"frmServerProperties.frx":1673
            Height          =   615
            Left            =   840
            TabIndex        =   30
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4935
         Index           =   2
         Left            =   -74760
         TabIndex        =   14
         Top             =   840
         Width           =   5655
         Begin VB.Frame Frame6 
            Caption         =   " Enable/Disable Logging "
            Height          =   2535
            Left            =   120
            TabIndex        =   15
            Top             =   1440
            Width           =   5415
            Begin VB.CheckBox EnableLogging 
               Caption         =   "Enable Event Logging"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   1935
            End
            Begin VB.ComboBox LogFormat 
               Height          =   315
               Left            =   1320
               TabIndex        =   16
               Top             =   1200
               Width           =   2895
            End
            Begin VB.Label Label6 
               Caption         =   "Check the box if you require event logging."
               Height          =   255
               Left            =   240
               TabIndex        =   20
               Top             =   720
               Width           =   3135
            End
            Begin VB.Label Label13 
               Caption         =   "Logfile Format:"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "Select the format you wish the logfile to be written in.  The default is the LocalWEB Format."
               Height          =   495
               Left            =   240
               TabIndex        =   18
               Top             =   1800
               Width           =   5055
            End
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   240
            Picture         =   "frmServerProperties.frx":1707
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label3 
            Caption         =   $"frmServerProperties.frx":1B49
            Height          =   855
            Left            =   840
            TabIndex        =   21
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5175
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   5775
         Begin VB.Frame Frame3 
            Height          =   2055
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   5415
            Begin VB.TextBox HomePage 
               Height          =   285
               Left            =   840
               TabIndex        =   12
               Top             =   1200
               Width           =   3435
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   120
               Picture         =   "frmServerProperties.frx":1C06
               Top             =   240
               Width           =   480
            End
            Begin VB.Label Label1 
               Caption         =   "Supply the location of where the pages that LocalWEB is serving can be found on this machine. For example: C:\WWW\HOMEPAGE"
               Height          =   615
               Left            =   720
               TabIndex        =   13
               Top             =   240
               Width           =   3735
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1695
            Left            =   120
            TabIndex        =   8
            Top             =   2760
            Width           =   5415
            Begin VB.TextBox LogFile 
               Height          =   285
               Left            =   840
               TabIndex        =   9
               Top             =   960
               Width           =   3555
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   120
               Picture         =   "frmServerProperties.frx":2048
               Top             =   240
               Width           =   480
            End
            Begin VB.Label Label2 
               Caption         =   "Supply the location of where you would like the log file generated by LocalWEB to be saved. For Example C:\WWW\LOGS"
               Height          =   615
               Left            =   840
               TabIndex        =   10
               Top             =   240
               Width           =   3615
            End
         End
      End
      Begin VB.Label Label19 
         Caption         =   $"frmServerProperties.frx":248A
         Height          =   1215
         Left            =   -74160
         TabIndex        =   73
         Top             =   840
         Width           =   5055
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmServerProperties.frx":25EE
         Top             =   840
         Width           =   480
      End
      Begin VB.Label ErrDesc 
         Caption         =   $"frmServerProperties.frx":2A30
         Height          =   975
         Left            =   -74160
         TabIndex        =   67
         Top             =   840
         Width           =   4815
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmServerProperties.frx":2B02
         Top             =   840
         Width           =   480
      End
      Begin VB.Image TNetImage 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmServerProperties.frx":2F44
         Top             =   840
         Width           =   480
      End
      Begin VB.Label TNetText 
         Caption         =   $"frmServerProperties.frx":3386
         Height          =   1215
         Left            =   -73920
         TabIndex        =   66
         Top             =   840
         Width           =   4815
      End
      Begin VB.Image DocImage 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmServerProperties.frx":348C
         Top             =   960
         Width           =   480
      End
      Begin VB.Label DocDescription 
         Caption         =   $"frmServerProperties.frx":38CE
         Height          =   855
         Left            =   -74160
         TabIndex        =   50
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label17 
         Caption         =   $"frmServerProperties.frx":39A6
         Height          =   855
         Left            =   -74160
         TabIndex        =   44
         Top             =   840
         Width           =   4695
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmServerProperties.frx":3A84
         Top             =   840
         Width           =   480
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServerProperties.frx":3EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServerProperties.frx":41E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServerProperties.frx":44FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServerProperties.frx":4814
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServerProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddDocument_Click()

    frmDocumentsAdd.Show 1
End Sub

Private Sub AllowCGI_Click()

    If AllowCGI.Value = 0 Then
      form1.CGISUPPORT.Caption = "Disabled"
    Else
      form1.CGISUPPORT.Caption = "Enabled"
    End If
    
End Sub

Private Sub Command1_Click()
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

    'write virtual dirs
    If VirtualDirectories.ListItems.Count > 0 Then
      FileHandle = FreeFile
      If FileExists(App.Path & "\vdirs.lst") Then
        Kill App.Path & "\vdirs.lst"
      End If
      Open App.Path & "\vdirs.lst" For Output As #FileHandle
      For f = 1 To VirtualDirectories.ListItems.Count
        Write #FileHandle, VirtualDirectories.ListItems.Item(f) & "," & VirtualDirectories.ListItems.Item(f).SubItems(1)
      Next f
      Close #FileHandle
    Else
      If FileExists(App.Path & "\vdirs.lst") Then
        Kill App.Path & "\vdirs.lst"
      End If
    End If
    
    'telnet server
    slValue = "TelnetEnabled"
    slData = TelnetEnabled.Value
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "TelnetPassword"
    slData = UCase(TelnetPassword.Text)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    'rewrite MIME type List
    FileHandle = FreeFile
    If FileExists(App.Path & "\mime.lst") Then
      Kill App.Path & "\mime.lst"
    End If
    Open App.Path & "\mime.lst" For Output As #FileHandle
    For f = 0 To MimeList.ListCount - 1
      Write #FileHandle, MimeList.List(f)
    Next f
    Close #FileHandle
    
    'documents and footer
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
    
    Rem ** Check IP Exception List, If items, save to EXCEPT.LST **
    
    
    If ExceptionList.ListItems.Count > 0 Then
      If FileExists(App.Path & "\except.lst") Then
         Kill (App.Path & "\except.lst")
      End If
      FileHandle = FreeFile
      Open App.Path & "\except.lst" For Output As #FileHandle
      Dim i
      For i = 1 To ExceptionList.ListItems.Count
        Write #FileHandle, ExceptionList.ListItems.Item(i)
      Next i
      Close #FileHandle
    Else
      Kill App.Path & "\except.lst"
    End If
    
    'write out error info to registry
    slValue = "Error404"
    slData = ErrorList.ListItems.Item(1).SubItems(1)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "CGINotFound"
    slData = ErrorList.ListItems.Item(2).SubItems(1)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "CGIUserBarred"
    slData = ErrorList.ListItems.Item(3).SubItems(1)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "CGIServerBarred"
    slData = ErrorList.ListItems.Item(4).SubItems(1)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    'write DNS4EVER registry values
    slValue = "DNS4EVEREnabled"
    slData = Str(DNS4everenabled.Value)
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "DNS4EVERDomainName"
    slData = dns4everdomainname.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "DNS4EVERUserName"
    slData = DNS4EVERUserName.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "DNS4EVERPassword"
    slData = DNS4EVERPassword.Text
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "DNS4EVERRefreshRate"
    slData = DNS4EVERRefreshRate.Text
    
    'Save IP Deny List
    If ServerDeny.ListItems.Count > 0 Then
      If FileExists(App.Path & "\deny.lst") Then
         Kill (App.Path & "\deny.lst")
      End If
      FileHandle = FreeFile
      Open App.Path & "\deny.lst" For Output As #FileHandle
      For i = 1 To ServerDeny.ListItems.Count
        Write #FileHandle, ServerDeny.ListItems.Item(i)
      Next i
      Close #FileHandle
    Else
      Kill App.Path & "\deny.lst"
    End If
    
    
    
    Unload Me
    
End Sub

Private Sub Command10_Click()
Dim SelectedDir As Integer
Dim itm As ListItem

    Load frmAddVDir
    frmAddVDir.Caption = "Edit Virtual Directory"
    SelectedDir = VirtualDirectories.SelectedItem.Index
    frmAddVDir.VirtualDir = VirtualDirectories.ListItems.Item(SelectedDir)
    frmAddVDir.PathName = VirtualDirectories.ListItems.Item(SelectedDir).SubItems(1)
    VirtualDirectories.ListItems.Remove (SelectedDir)
    frmAddVDir.Show 1
    
    
End Sub

Private Sub Command11_Click()
Dim SelectedError As Integer
    
    
    SelectedError = ErrorList.SelectedItem.Index
    CommonDialog1.FileName = ErrorList.ListItems.Item(SelectedError).SubItems(1)
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "All Files (*.*)|*.*|HTML Files(*.htm)|*.htm|HTML Files(*.html)|*.html"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    ErrorList.ListItems.Item(SelectedError).SubItems(1) = CommonDialog1.FileName
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Command12_Click()

    frmSecurityIP.Show 1
    
End Sub

Private Sub Command13_Click()
Dim SelectedIP As Integer

    If ServerDeny.ListItems.Count > 0 Then
      SelectedIP = ServerDeny.SelectedItem.Index
      ServerDeny.ListItems.Remove (SelectedIP)
    End If
End Sub

Private Sub Command2_Click()
Dim SelectedError As Integer
    
    
    SelectedError = ErrorList.SelectedItem.Index
    ErrorList.ListItems.Item(SelectedError).SubItems(1) = "Default"
End Sub

Private Sub Command3_Click()

    Unload Me
    
End Sub

Private Sub Command4_Click()
   
    frmAddMIME.Show 1
End Sub

Private Sub Command5_Click()

    frmAddVDir.Show 1
End Sub

Private Sub Command6_Click()

    frmIPAddress.Show 1
End Sub

Private Sub Command7_Click()
Dim SelectedIP As Integer

    If ExceptionList.ListItems.Count > 0 Then
      SelectedIP = ExceptionList.SelectedItem.Index
      ExceptionList.ListItems.Remove (SelectedIP)
    End If
End Sub

Private Sub Command8_Click()
Dim Response
    
    Response = MsgBox("Are you sure you want to remove that MIME Header?", vbQuestion + vbYesNo, "LocalWEB")
    If Response = vbYes Then
      MimeList.RemoveItem (MimeList.ListIndex)
    End If
End Sub

Private Sub Command9_Click()
Dim SelectedDir As Integer

    If VirtualDirectories.ListItems.Count > 0 Then
      SelectedDir = VirtualDirectories.SelectedItem.Index
      VirtualDirectories.ListItems.Remove (SelectedDir)
    End If
    
    
    
End Sub

Private Sub DNS4EVEREnabled_Click()

     If DNS4everenabled.Value = 0 Then
        dns4everdomainname.Enabled = False
        DNS4EVERUserName.Enabled = False
        DNS4EVERPassword.Enabled = False
        DNS4EVERRefreshRate.Enabled = False
        form1.DNS4everenabled.Caption = "Disabled"
        form1.Command2.Enabled = False
        form1.DNS4everupdate.Enabled = False
     Else
        dns4everdomainname.Enabled = True
        DNS4EVERUserName.Enabled = True
        DNS4EVERPassword.Enabled = True
        DNS4EVERRefreshRate.Enabled = True
        form1.DNS4everenabled.Caption = "Enabled"
        form1.Command2.Enabled = True
        form1.DNS4everupdate.Enabled = True
     End If
End Sub

Private Sub Form_Load()
Dim FileHandle
Dim Header As String
Dim itm As ListItem
Dim FName As String
Dim VDirs As String
Dim spointer As Integer

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
        Set itm = ExceptionList.ListItems.Add(, , IPAddress, , 3)
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
        MimeList.AddItem Header
      Wend
      Close #FileHandle
    End If
    
    'Documents Tab Data
    'Get Page Footer Information
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
      
      'populate virtual directory list
      If FileExists(App.Path & "\vdirs.lst") Then
        FileHandle = FreeFile
        Dim VirtualDir, RealDir As String
        Open App.Path & "\vdirs.lst" For Input As #FileHandle
        Do While Not EOF(FileHandle)
          RealDir = ""
          VitrualDir = ""
          Input #FileHandle, VDirs
          spointer = InStr(1, VDirs, ",")
          VirtualDir = Left(VDirs, spointer - 1)
          For f = spointer + 1 To Len(VDirs)
            RealDir = RealDir & Mid(VDirs, f, 1)
          Next f
          Set itm = VirtualDirectories.ListItems.Add(, , VirtualDir)
          itm.SubItems(1) = RealDir
        Loop
        Close #FileHandle
      End If
      
      'telnet server
      TelnetEnabled.Value = GetRegValue(hKey, SubKey, "TelnetEnabled", "1")
      TelnetPassword = GetRegValue(hKey, SubKey, "TelnetPassword", "PASSWORD")
    
      'check if telnet server is enabled
      If TelnetEnabled.Value = 0 Then
        TelnetPassword.Enabled = False
      Else
        TelnetPassword.Enabled = True
      End If
      
      'populate error list
      Set itm = ErrorList.ListItems.Add(, , "404 - Not Found")
      Set itm = ErrorList.ListItems.Add(, , "404.1 - CGI Script Not Found")
      Set itm = ErrorList.ListItems.Add(, , "404.2 - CGI Barred on Your Machine")
      Set itm = ErrorList.ListItems.Add(, , "404.3 - CGI Barred on This Server")
      ErrorList.ListItems.Item(1).SubItems(1) = GetRegValue(hKey, SubKey, "Error404", "Default")
      ErrorList.ListItems.Item(2).SubItems(1) = GetRegValue(hKey, SubKey, "CGINotFound", "Default")
      ErrorList.ListItems.Item(3).SubItems(1) = GetRegValue(hKey, SubKey, "CGIUserBarred", "Default")
      ErrorList.ListItems.Item(4).SubItems(1) = GetRegValue(hKey, SubKey, "CGIServerBarred", "Default")
      
      'setup DNS4EVER registry setting
      DNS4everenabled.Value = Val(GetRegValue(hKey, SubKey, "DNS4EVEREnabled", "0"))
      dns4everdomainname = GetRegValue(hKey, SubKey, "DNS4EVERDomainName", "<not set>")
      DNS4EVERUserName = GetRegValue(hKey, SubKey, "DNS4EVERUsername", "<not set>")
      DNS4EVERPassword = GetRegValue(hKey, SubKey, "DNS4EVERPassword", "<not set>")
      DNS4EVERRefreshRate = GetRegValue(hKey, SubKey, "DNS4EVERRefreshRate", "60")
      If DNS4everenabled.Value = 0 Then
        dns4everdomainname.Enabled = False
        DNS4EVERUserName.Enabled = False
        DNS4EVERPassword.Enabled = False
        DNS4EVERRefreshRate.Enabled = False
      End If
      
      'setup Sever Security
      FileHandle = FreeFile
      If FileExists(App.Path & "\deny.lst") Then
        Open App.Path & "\deny.lst" For Input As #FileHandle
        Dim IPAddressBlock As String
        While Not EOF(FileHandle)
          Input #FileHandle, IPAddressBlock
          Set itm = ServerDeny.ListItems.Add(, , IPAddressBlock, , 3)
        Wend
        Close #FileHandle
      End If
        
End Sub

Private Sub MoveDown_Click()
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

Private Sub MoveUp_Click()
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

Private Sub RemoveDocument_Click()
Dim Response
    
    Response = MsgBox("Are you sure you want to remove that Default Document?", vbQuestion + vbYesNo, "LocalWEB")
    If Response = vbYes Then
      DefaultList.RemoveItem (DefaultList.ListIndex)
    End If
End Sub

Private Sub TelnetEnabled_Click()

    If TelnetEnabled.Value = 1 Then
      TelnetPassword.Enabled = True
      form1.Telnetsupport.Caption = "Enabled"
    Else
      TelnetPassword.Enabled = False
      form1.Telnetsupport.Caption = "Disabled"
    End If
    
End Sub
