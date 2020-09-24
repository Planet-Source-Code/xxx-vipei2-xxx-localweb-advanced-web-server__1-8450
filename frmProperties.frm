VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LocalWEB Properties"
   ClientHeight    =   5805
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6285
   Icon            =   "frmProperties.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog LogFileSelector 
      Left            =   600
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSComDlg.CommonDialog HomePageSelector 
      Left            =   120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8493
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "File Locations"
      TabPicture(0)   =   "frmProperties.frx":000C
      Tab(0).ControlCount=   7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TxtHomePage"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TxtLogFile"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      TabCaption(1)   =   "Server Options"
      TabPicture(1)   =   "frmProperties.frx":0028
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "CGI Scripts"
      TabPicture(2)   =   "frmProperties.frx":0044
      Tab(2).ControlCount=   0
      Tab(2).ControlEnabled=   0   'False
      Begin VB.Frame Frame4 
         Caption         =   "Server Settings"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   5655
         Begin VB.CheckBox Check1 
            Caption         =   "Enable Connection Logging"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Frame Frame5 
            Caption         =   "Connection Logging"
            Height          =   1455
            Left            =   120
            TabIndex        =   24
            Top             =   2400
            Width           =   5415
            Begin VB.CheckBox Check6 
               Caption         =   "CGI Scripts Executed"
               Height          =   255
               Left            =   2520
               TabIndex        =   29
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Images Downloaded"
               Height          =   255
               Left            =   2520
               TabIndex        =   28
               Top             =   360
               Width           =   1935
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Pages Requested"
               Height          =   255
               Left            =   240
               TabIndex        =   27
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Date and Time"
               Height          =   255
               Left            =   240
               TabIndex        =   26
               Top             =   720
               Width           =   1455
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Remote IP Address"
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   360
               Width           =   1815
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Connection Settings"
            Height          =   1455
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   5415
            Begin VB.TextBox TxtPort 
               Height          =   285
               Left            =   1080
               TabIndex        =   20
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label4 
               Caption         =   "Port"
               Height          =   255
               Left            =   240
               TabIndex        =   23
               Top             =   720
               Width           =   375
            End
            Begin VB.Label LocalIPAddress 
               Caption         =   "Label6"
               ForeColor       =   &H0080FFFF&
               Height          =   255
               Left            =   1080
               TabIndex        =   22
               Top             =   360
               Width           =   4095
            End
            Begin VB.Label Label5 
               Caption         =   "IP Address"
               Height          =   255
               Left            =   240
               TabIndex        =   21
               Top             =   360
               Width           =   855
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   360
         TabIndex        =   16
         Top             =   2640
         Width           =   5055
         Begin VB.Label Label3 
            Caption         =   $"frmProperties.frx":0060
            ForeColor       =   &H00FF0000&
            Height          =   855
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   5040
         TabIndex        =   2
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox TxtLogFile 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox TxtHomePage 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Log File"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Home Page"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   11
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   10
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    MsgBox "Place code here to set options w/o closing dialog!"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    MsgBox "Place code here to set options and close dialog!"
    Unload Me
End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    LocalIPAddress.Caption = Form1.Winsock1.LocalIP
End Sub

Private Sub tbsOptions_Click()

    
End Sub

