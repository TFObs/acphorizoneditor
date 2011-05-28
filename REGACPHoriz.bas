Attribute VB_Name = "REGACPHoriz"
Option Explicit
'zunächst die benötigten Deklarationen
Private Declare Function GetSystemMenu Lib "user32" _
        (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" _
        (ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" _
        (ByVal hwnd As Long) As Long
  
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal _
        lpSubKey As String, ByVal ulOptions As Long, ByVal _
        samDesired As Long, phkResult As Long) As Long
        
Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
        
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Any) As Long
        Private Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, lpData As Long, ByVal cbData As Long) _
        As Long
        
Private Declare Function RegSetValueEx_Str Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, ByVal lpData As String, ByVal cbData As _
        Long) As Long

        
Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_USERS As Long = &H80000003
Private Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Private Const HKEY_CURRENT_CONFIG As Long = &H80000005
Private Const HKEY_DYN_DATA As Long = &H80000006


Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20

Private Const KEY_READ As Long = KEY_QUERY_VALUE Or _
                 KEY_ENUMERATE_SUB_KEYS _
                 Or KEY_NOTIFY
                 
Private Const KEY_ALL_ACCESS As Long = KEY_QUERY_VALUE Or _
                       KEY_SET_VALUE Or _
                       KEY_CREATE_SUB_KEY Or _
                       KEY_ENUMERATE_SUB_KEYS Or _
                       KEY_NOTIFY Or _
                       KEY_CREATE_LINK
                       
Private Const ERROR_SUCCESS As Long = 0&

Private Const REG_NONE As Long = 0&
Private Const REG_SZ As Long = 1&
Private Const REG_EXPAND_SZ As Long = 2&
Private Const REG_BINARY As Long = 3&
Private Const REG_DWORD As Long = 4&
Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4&
Private Const REG_DWORD_BIG_ENDIAN As Long = 5&
Private Const REG_LINK As Long = 6&
Private Const REG_MULTI_SZ As Long = 7&

Private Const REG_OPTION_NON_VOLATILE As Long = &H0&


Private Const SC_CLOSE = &HF060
Private Const MF_BYCOMMAND = &H0

'Declaration for the Type to help debugging
Public Type DebugHelper
    module As String
    place As String
End Type


'Public scope As acp.Telescope'<<<--DON'T DO This: ACP 6.0 changed ActiveX interfaces: No early binding!!!
Public scope As Object 'New way to reference ACP.Telescope
Public abort As Boolean
Public x As Integer


'Removal of the "Close"-Button
Public Sub DisableCloseButton(hwnd As Long)
  Dim hMenu As Long
  hMenu = GetSystemMenu(hwnd, 0&)
  If hMenu Then
    Call DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    DrawMenuBar hwnd
  End If
End Sub


Public Sub getACPHorizon()
Dim result&, value As Variant
Dim azimval
      
    'Read String-Value
    result = RegValueGet(HKEY_LOCAL_MACHINE, "Software\denny\acp\Observatory", "Horizon", value)
 
    azimval = Split(value, " ")
    
    For x = 0 To UBound(azimval) - 1
        frmMain.GridHoriz.TextMatrix(x + 1, 1) = azimval(x)
    Next x
    
    frmMain.DrawHorizon
End Sub

Public Sub saveACPHorizon()
Dim result&, value As String
Dim wmi, answer, wql, acp
FloatWindow frmMain.hwnd, False
Set wmi = GetObject("winmgmts:")

' create request
wql = "select * from win32_process where name='acp.exe'"

' Send
Set answer = wmi.ExecQuery(wql)

If answer.Count > 0 Then
FloatWindow frmMain.hwnd, False
    result = MsgBox("ACP is already running!" & vbCrLf & "ACP has to be closed for this operation" & _
    "to be successful!" & vbCrLf & vbCrLf & "clicking OK will TERMINATE ACP" & vbCrLf & vbCrLf & _
    "Please Disconnect Scope/Dome,Camera and Weather before clicking the OK-Button...", vbOKCancel + vbExclamation, "Save Horizon")
   
   If result = 1 Then
   
    On Error Resume Next
    scope.Connected = False
    abort = True
    Err.Clear
   
        For Each acp In answer
            acp.Terminate 0
        Next
        
    Else:  Exit Sub
    
    End If
    
End If

        value = ""
        For x = 1 To 180
            If IsNumeric(frmMain.GridHoriz.TextMatrix(x, 1)) Then
                value = value & frmMain.GridHoriz.TextMatrix(x, 1) & " "
            Else: value = value & "0.0 "
            End If
        Next x
    
        value = Trim(value)
    
        'Write String-Value
        result = RegValueSet(HKEY_LOCAL_MACHINE, "Software\denny\acp\Observatory", "Horizon", value)
        
        If result = 0 Then
            MsgBox "Data successfully written!", vbInformation, "Write successful!"
        Else
            MsgBox "Error writing the Data, please write to file and try again!", vbCritical, "Error"
        End If
  FloatWindow frmMain.hwnd, True
End Sub


Function RegValueGet(Root&, Key$, Field$, value As Variant) As Long
  Dim result&, hKey&, dwType&, Lng&, Buffer$, l&
    'Read Value from Reg-Field
    result = RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey)
    If result = ERROR_SUCCESS Then
      result = RegQueryValueEx(hKey, Field, 0&, dwType, ByVal 0&, l)
      If result = ERROR_SUCCESS Then
        Select Case dwType
          Case REG_SZ
            Buffer = Space$(l + 1)
            result = RegQueryValueEx(hKey, Field, 0&, _
                                     dwType, ByVal Buffer, l)
            If result = ERROR_SUCCESS Then value = Buffer
          Case REG_DWORD
            result = RegQueryValueEx(hKey, Field, 0&, dwType, Lng, l)
            If result = ERROR_SUCCESS Then value = Lng
        End Select
      End If
    End If
    
    If result = ERROR_SUCCESS Then result = RegCloseKey(hKey)
    RegValueGet = result
End Function

Function RegValueSet(Root&, Key$, Field$, value As Variant) As Long
  Dim result&, hKey&, s$, l&
    'Write Value to Regfield
    result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, hKey)
    If result = ERROR_SUCCESS Then
      Select Case VarType(value)
        Case vbInteger, vbLong
          l = CLng(value)
          result = RegSetValueEx(hKey, Field, 0, REG_DWORD, l, 4)
        Case vbString
          s = CStr(value)
          result = RegSetValueEx_Str(hKey, Field, 0, REG_SZ, s, _
                                        Len(s) + 1)
      End Select
      result = RegCloseKey(hKey)
    End If
    
    RegValueSet = result
End Function

