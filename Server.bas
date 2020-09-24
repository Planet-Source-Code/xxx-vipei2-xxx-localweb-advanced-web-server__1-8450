Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverLapped As Any) As Long


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&

      
'Added by STF
'For the System Tray
      'user defined type required by Shell_NotifyIcon API call
      Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type

      'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

      Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Public nid As NOTIFYICONDATA
 'End STF Add
 



Sub ServeImages(FileName As String, Index As Integer)
Rem *************************************************
Rem **** THIS CODE IS NOW NOT USED Phil 23/12/99  ***
Rem *************************************************
Dim FLoc As Long
'Dim szBuffer As String
Dim szBuffer() As Byte
Dim lBlockSize As Long
Dim Counter, Remainder As Long
Dim BServed As Long

Dim barray() As Byte
Dim temparray() As Byte
Dim i, i1 As Long
Dim FreeHandle

    FreeHandle = FreeFile
    Open FileName For Binary Access Read As #FreeHandle
    'BServed = Val(Form1.BytesServed.Caption)
    'BServed = BServed + FileLen(FileName)
    'Form1.BytesServed.Caption = BServed
    'Do While FLoc < LOF(1)
      'szBuffer = Space$(1024)
     ' ReDim szBuffer(1024)
      'Get #1, , szBuffer
      'Form1.Winsock1.SendData szBuffer
      'FLoc = Loc(1)
    'Loop
    ReDim barray(LOF(FreeHandle))
    i = 0
    Do While Not EOF(FreeHandle)
      Get #FreeHandle, , barray(i)
      i = i + 1
    Loop
    
    Rem *** check if the buffer is over 8Kb ***
    Rem *** if So need to break down into chunks as Winsock will only send max 8K chunks ***
    
    If LOF(FreeHandle) > 8192 Then
      'Send buffer in 1K chunks
      Counter = LOF(FreeHandle) \ 8192
      Remainder = LOF(FreeHandle) Mod 8192
      Do
       For i = 0 To 8192
         Form1.Winsock1(Index).SendData barray(i + i1)
         'temparray(i) = barray(i + i1)
       Next i
       i1 = i1 + 8192
      
       'Form1.Winsock1(Index).SendData temparray
       Counter = Counter - 1
      
      Loop Until Counter = 0
      
      'nasty bit we have to do not to lose Winsock events
      Do Until Form1.Winsock1(Index).State = sckConnected
        DoEvents
      Loop
      
      'now send remainder of bytes, if any
            
      For i = 1 To Remainder
        Form1.Winsock1(Index).SendData barray(i + i1)
      Next i
    Else
      Form1.Winsock1(Index).SendData barray
    End If
    
    Close #FreeHandle
End Sub

Sub ServeText(FileName As String, Index As Integer)
Rem *************************************************
Rem **** THIS CODE IS NOW NOT USED Phil 23/12/99  ***
Rem *************************************************
Dim TextLine As String
Dim FreeHandle

    FreeHandle = FreeFile
    Open FileName For Input As #FreeHandle
    While Not EOF(FreeHandle)
      Line Input #FreeHandle, TextLine
      Form1.Winsock1(Index).SendData TextLine
    Wend
    Close #FreeHandle
    DoEvents
End Sub
Sub ExecuteCGI(CommandLine As String, WebROOT As String, Index As Integer)
Dim ConvertedLine As String
Dim PathAndScript As String
Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
Dim start As STARTUPINFO
Dim SA As SECURITY_ATTRIBUTES, hReadPipe As Long, hWritePipe As Long
Dim bytesread As Long, mybuff As String
Dim i As Integer
Dim f As Integer
Dim ErrorName As String

    
    Rem *** Check if CGI execution is enabled ***
    
    
    DoEvents
    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\LocalWEB\Settings"
    If Val(GetRegValue(hKey, SubKey, "AllowCGI", "0")) = 0 Then
      If GetRegValue(hKey, SubKey, "CGIServerBarred", "Default") = "Default" Then
         Form1.Winsock1(Index).SendData "HTTP/1.0 200 OK"
         Form1.Winsock1(Index).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
         Form1.Winsock1(Index).SendData "<HTML>"
         Form1.Winsock1(Index).SendData "<BODY>"
         Form1.Winsock1(Index).SendData "<TITLE>CGI Warning</TITLE>"
         Form1.Winsock1(Index).SendData "<h1>CGI Warning</h1>"
         Form1.Winsock1(Index).SendData "<p>CGI execution has been disabled on this server.</p>" + vbCrLf + vbCrLf
         Form1.Winsock1(Index).SendData "<p><em>LocalWEB HTTP Server</em></p>"
         Form1.Winsock1(Index).SendData "</BODY>"
         Form1.Winsock1(Index).SendData "</HTML>"
         Exit Sub
      Else
         ErrorName = GetRegValue(hKey, SubKey, "CGIServerBarred", "Default")
         Form1.Winsock1(Index).SendData "HTTP/1.0 200 OK"
         Form1.Winsock1(Index).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
         ServeText ErrorName, Index
         Exit Sub
      End If
    End If
    
    Rem ** if Here CGI is enabled **
    
    Rem ** Now need to see if requesting machine is allowed to run CGI **
    
    If FileExists(App.Path & "\except.lst") Then
      Dim FileHandle
      Dim ListAddress As String
      FileHandle = FreeFile
      Open App.Path & "\except.lst" For Input As #FileHandle
      While Not EOF(FileHandle)
        Input #FileHandle, ListAddress
        If ListAddress = Form1.Winsock1(Index).RemoteHostIP Then
          If GetRegValue(hKey, SubKey, "CGIUserBarred", "Default") = "Default" Then
           Form1.Winsock1(Index).SendData "HTTP/1.0 200 OK"
           Form1.Winsock1(Index).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
           Form1.Winsock1(Index).SendData "<HTML>"
           Form1.Winsock1(Index).SendData "<BODY>"
           Form1.Winsock1(Index).SendData "<TITLE>CGI Warning</TITLE>"
           Form1.Winsock1(Index).SendData "<h1>CGI Warning</h1>"
           Form1.Winsock1(Index).SendData "<p>Your machine has been blocked from executing CGI Scripts.</p>" + vbCrLf + vbCrLf
           Form1.Winsock1(Index).SendData "<p><em>LocalWEB HTTP Server</em></p>"
           Form1.Winsock1(Index).SendData "</BODY>"
           Form1.Winsock1(Index).SendData "</HTML>"
           Close #FileHandle
           Exit Sub
          Else
           ErrorName = GetRegValue(hKey, SubKey, "CGIUserBarred", "Default")
           Form1.Winsock1(Index).SendData "HTTP/1.0 200 OK"
           Form1.Winsock1(Index).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
           ServeText ErrorName, Index
           Exit Sub
          End If
        End If
      Wend
      Close #FileHandle
    End If
      
    Rem ******************************************************************
    Rem ** If we are here CGI is enabled, and the client is not blocked **
    Rem ******************************************************************
    
    Rem ** Now we need to check for arguments to pass to the CGI Script **
    
    Dim ArgumentStart
    Dim Arguments As String
    Dim ScriptName As String
    Dim OutputName As String
    
    ArgumentStart = InStr(CommandLine, "?")
    If ArgumentStart <> 0 Then
      'OK - we have Arguments create interface file
      For f = ArgumentStart + 1 To Len(CommandLine)
        Arguments = Arguments + Mid(CommandLine, f, 1)
      Next f
      For f = 1 To ArgumentStart - 1
        ScriptName = ScriptName + Mid(CommandLine, f, 1)
      Next f
      'Create Input file for CGI Script
      FileHandle = FreeFile
      OutputName = "c:\cgi" & FileHandle & ".ini"
      Open OutputName For Output As #FreeFile
      Print #FileHandle, "[Environment]"
      Print #FileHandle, "QUERY_STRING=" + Arguments
      Print #FileHandle, "CONTENT_LENGTH=" & Str(Len(Arguments))
      Print #FileHandle, "GATEWAY_INTERFACE=CGI/1.1"
      Print #FileHandle, "REQUEST_METHOD=GET"
      Print #FileHandle, "SERVER_NAME=LocalWEB"
      Print #FileHandle, "SERVER_SOFTWARE=West Street LocalWEB"
      Print #FileHandle, "SERVER_PORT=80"
      Print #FileHandle, "SCRIPT_NAME=" & ScriptName
      Print #FileHandle, "PATH_INFO=" & WebROOT
      Print #FileHandle, "REMOTE_ADDR=" & Form1.Winsock1(Index).RemoteHostIP
      Close #FileHandle
    Else
      'There are no arguments
      'Create Input file for CGI Script
      FileHandle = FreeFile
      ScriptName = CommandLine
      OutputName = "c:\cgi" & FileHandle & ".ini"
      Open OutputName For Output As #FileHandle
      Print #FileHandle, "[Environment]"
      Print #FileHandle, "QUERY_STRING=None"
      Print #FileHandle, "CONTENT_LENGTH=0"
      Print #FileHandle, "GATEWAY_INTERFACE=CGI/1.1"
      Print #FileHandle, "REQUEST_METHOD=GET"
      Print #FileHandle, "SERVER_NAME=LocalWEB"
      Print #FileHandle, "SERVER_SOFTWARE=West Street LocalWEB"
      Print #FileHandle, "SERVER_PORT=80"
      Print #FileHandle, "SCRIPT_NAME=" & ScriptName
      Print #FileHandle, "PATH_INFO=" & WebROOT
      Print #FileHandle, "REMOTE_ADDR=" & Form1.Winsock1(Index).RemoteHostIP
      Close #FileHandle
    End If
      
    Rem ** Now Execute the Script (CGI Arguments have been stripped **
    
    For f = 1 To Len(ScriptName)
      If Mid$(ScriptName, f, 1) = "/" Then
        ConvertedLine = ConvertedLine & "\"
      Else
        ConvertedLine = ConvertedLine & Mid$(ScriptName, f, 1)
      End If
    Next f
    
    
    PathAndScript = WebROOT + "\" + ConvertedLine + ".exe"
    
    Rem ** Now lets see if the CGI script exists, if so run it **
    
    If Not FileExists(PathAndScript) Then
      If GetRegValue(hKey, SubKey, "CGINotFound", "Default") = "Default" Then
        Form1.Winsock1(Index).SendData "HTTP/1.0 200 OK"
        Form1.Winsock1(Index).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
        Form1.Winsock1(Index).SendData "<HTML>"
        Form1.Winsock1(Index).SendData "<BODY>"
        Form1.Winsock1(Index).SendData "<TITLE>CGI Warning</TITLE>"
        Form1.Winsock1(Index).SendData "<h1>CGI Warning</h1>"
        Form1.Winsock1(Index).SendData "<p>CGI script could not be run -- it does not exist on the server.</p>" + vbCrLf + vbCrLf
        Form1.Winsock1(Index).SendData "<p><em>LocalWEB HTTP Server</em></p>"
        Form1.Winsock1(Index).SendData "</BODY>"
        Form1.Winsock1(Index).SendData "</HTML>"
        Exit Sub
      Else
        ErrorName = GetRegValue(hKey, SubKey, "CGINotFound", "Default")
        Form1.Winsock1(Index).SendData "HTTP/1.0 200 OK"
        Form1.Winsock1(Index).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
        ServeText ErrorName, Index
        Exit Sub
      End If
    End If
    
    mybuff = String(8192, Chr$(65))
    
    SA.nLength = Len(SA)
    SA.bInheritHandle = 1&
    SA.lpSecurityDescriptor = 0&
    
    ret = CreatePipe(hReadPipe, hWritePipe, SA, 0)
    If ret = 0 Then
      MsgBox "CreatePipe failed: Error: " & Err.LastDllError
      Exit Sub
    End If
    
    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES
    start.hStdOutput = hWritePipe
    
    'If ArgumentStart <> 0 Then
      PathAndScript = PathAndScript + " " + OutputName
    'End If
       
    ret& = CreateProcessA(0&, PathAndScript, SA, SA, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    
    
    If ret <> 1 Then
      MsgBox "CreateProcess failed: Error: " & Err.LastDllError
    End If
    
    bSuccess = ReadFile(hReadPipe, mybuff, 1024, bytesread, 0&)
    If bSuccess = 1 Then
      'Text1.Text = Left(mybuff, bytesread)
      'Form1.StatsBox.AddItem Left(mybuff, bytesread)
      Form1.Winsock1(Index).SendData "HTTP/1.0 200 OK"
      Form1.Winsock1(Index).SendData "Content-Type: text/html" & vbCrLf & vbCrLf
      Form1.Winsock1(Index).SendData Left(mybuff, bytesread)
    Else
      MsgBox "ReadFile failed: Error: " & Err.LastDllError
    End If
    ret& = WaitForSingleObject(proc.hProcess, -1&)
    ret& = CloseHandle(proc.hProcess)
    ret& = CloseHandle(proc.hThread)
    ret& = CloseHandle(hReadPipe)
    ret& = CloseHandle(hWritePipe)
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

Sub BuildNewMimeList()
Dim FreeHandle

    FreeHandle = FreeFile
    Open App.Path & "\mime.lst" For Output As #FreeHandle
    Write #FreeHandle, "htm,text/html"
    Write #FreeHandle, "html,text/html"
    Write #FreeHandle, "jpg,image/jpeg"
    Write #FreeHandle, "jpeg,image/jpeg"
    Write #FreeHandle, "gif,image/gif"
    Write #FreeHandle, "txt,text/plain"
    Write #FreeHandle, "text,text/plain"
    Write #FreeHandle, "wav,audio,x-wav"
    Close #FreeHandle
    
    
    
End Sub

Sub BuildNewDefaultList()
Dim FreeHandle

    FreeHandle = FreeFile
    Open App.Path & "\default.lst" For Output As #FreeHandle
    Write #FreeHandle, "index.htm"
    Write #FreeHandle, "index.html"
    Write #FreeHandle, "default.htm"
    Write #FreeHandle, "default.html"
    Close #FreeHandle
    
End Sub
