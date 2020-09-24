Attribute VB_Name = "RegModule"
'Global Variables
Public glRegistered As Boolean
Public glRegNumber As String
Public SubKey As String
Public hKey As Long
Public Create As Long

'Registry Constants
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

'Registry Specific Access Rights
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = &H3F

'Open/Create Options
Public Const REG_OPTION_NON_VOLATILE = 0&
Public Const REG_OPTION_VOLATILE = &H1

'Key creation/open disposition
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

'masks for the predefined standard access types
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF

'Define severity codes
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_ACCESS_DENIED = 5
Public Const ERROR_NO_MORE_ITEMS = 259

'Predefined Value Types
'Structures Needed For Registry Prototypes
Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type


'Registry Function Prototypes
Declare Function RegOpenKeyEx Lib "advapi32" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey _
    As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long

Declare Function RegSetValueEx Lib "advapi32" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal _
    lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, ByVal szData As String, _
    ByVal cbData As Long) As Long

Declare Function RegCloseKey Lib "advapi32" _
    (ByVal hKey As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal _
    lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, ByVal szData As String, _
    ByRef lpcbData As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal _
    lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As _
    Long, ByVal samDesired As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    phkResult As Long, lpdwDisposition As Long) As Long

Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
    "RegDeleteKeyA" (ByVal hKey As Long, ByVal _
    lpSubKey As String) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey As Long, ByVal _
    lpValueName As String) As Long

'------------------------------------------------------------
'Changes or sets the Value
'------------------------------------------------------------
Function SetRegValue(hKey As Long, lpszSubKey As String, _
                    ByVal sSetValue As String, _
                    ByVal sValue As String) As Boolean

    On Error GoTo ErrorRoutineErr:
    
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    
    'Note: This function will create the key or
    'value if it doesn't exist.
    'Open or Create the key
    RegCreateKeyEx hKey, lpszSubKey, 0, "", _
        REG_OPTION_NON_VOLATILE, _
        KEY_ALL_ACCESS, SA, phkResult, Create
    
    lResult = RegSetValueEx(phkResult, sSetValue, 0, _
        REG_SZ, sValue, _
        CLng(Len(sValue) + 1))
    
    'Close the key
    RegCloseKey phkResult
    
    'Return SetRegValue Result
    SetRegValue = (lResult = ERROR_SUCCESS)
    Exit Function

ErrorRoutineErr::
  MsgBox "ERROR #" & Str$(Err) & " : " & Error & Chr(13) _
         & "Please exit and try again."
  SetRegValue = False

End Function

'------------------------------------------------------------
'Get the value of the key
'------------------------------------------------------------
Function GetRegValue(hKey As Long, lpszSubKey As String, _
    szKey As String, szDefault As String) As Variant

    On Error GoTo ErrorRoutineErr:
    
    Dim phkResult As Long
    Dim lResult As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    
    'Create Buffer
    szBuffer = Space(255)
    lBuffSize = Len(szBuffer)
    
    'Open the key
    RegOpenKeyEx hKey, lpszSubKey, 0, 1, phkResult
    
    'Query the value
    lResult = RegQueryValueEx(phkResult, szKey, 0, _
        0, szBuffer, lBuffSize)
    
    'Close the key
    RegCloseKey phkResult
    
    'Return obtained value
    If lResult = ERROR_SUCCESS Then
        GetRegValue = Left(szBuffer, lBuffSize - 1)
    Else
        GetRegValue = szDefault
    End If
    Exit Function
    
ErrorRoutineErr::
    MsgBox "ERROR #" & Str$(Err) & " : " & Error & Chr(13) _
         & "Please exit and try again."
    GetRegValue = ""

End Function




'------------------------------------------------------------
'Create the new key
'------------------------------------------------------------
Function CreateRegKey(NewSubKey As String) As Boolean

    On Error GoTo ErrorRoutineErr:
    
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    
    'Create key if it does not exist
    CreateRegKey = (RegCreateKeyEx(hKey, NewSubKey, _
        0, "", REG_OPTION_NON_VOLATILE, _
        KEY_ALL_ACCESS, SA, phkResult, Create) = ERROR_SUCCESS)
    
    'Close the key
    RegCloseKey phkResult
    Exit Function
    
ErrorRoutineErr::
    MsgBox "ERROR #" & Str$(Err) & " : " & Error & Chr(13) _
         & "Please exit and try again."
    CreateRegKey = False
  
End Function

Sub DisplayStep(StepNumber As Integer, LastStep As Integer)

    If StepNumber = 0 Then
      frmWizard.Back.Enabled = False
    Else
      frmWizard.Back.Enabled = True
    End If
    If StepNumber = LastStep Then
      frmWizard.Next.Enabled = False
      frmWizard.Finish.Enabled = True
    Else
      frmWizard.Next.Enabled = True
      frmWizard.Finish.Enabled = False
    End If
    
    frmWizard.Frame1(StepNumber).Left = 3240
    frmWizard.Frame1(StepNumber).Top = 120
    frmWizard.Frame1(StepNumber).Visible = True
    
End Sub

Function CheckRegistry() As Boolean
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
    
    'check if there is any reg info for LocalWEB
    
    If RegOpenKeyEx(hKey, SubKey, 0, 1, phkResult) = ERROR_SUCCESS Then
      CheckRegistry = True
    Else
      CheckRegistry = False
    End If
    
    'If CreateRegKey(SubKey) Then
    '  rc = SetRegValue(hKey, SubKey, slValue, slData)
    'Else
    '  MsgBox "Cannot create registry value"
    'End If
End Function

Sub CreateRegKeys()
Dim rc
Dim slValue, slData As String

    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\LocalWEB\Settings"
    
    slValue = "HomePage"
    slData = App.Path
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "LogFile"
    slData = App.Path
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "EnableEvents"
    slData = "1"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "LogFormat"
    slData = "LocalWEB Format"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "AllowCGI"
    slData = "0"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "TipOfTheDay"
    slData = "1"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    slValue = "CGINotFound"
    slData = "Default"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "CGIServerBarred"
    slData = "Default"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "CGIUserBarred"
    slData = "Default"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "Error404"
    slData = "Default"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "FooterEnabled"
    slData = "0"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    slValue = "FooterCode"
    slData = "<p><font color=&quot;#0000FF&quot; size=&quot;1&quot;>This site is served by LocalWEB</font></p>"
    rc = SetRegValue(hKey, SubKey, slValue, slData)
    
    
End Sub
