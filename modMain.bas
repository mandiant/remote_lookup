Attribute VB_Name = "modMain"
'Contributed by FireEye FLARE Team
'Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
'Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
'License: GPL


'This will handle the installation of the ActX controls if it is not already found on the system
'this runs from a module before the main form is loaded because the main form depends on these.
'I did not want to use a batch file because there is a strong chance that newer versions of Windows
'which are 64-bit would try to execute it with a 64-bit cmd.exe which would then screw up regsvr32 call.

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'vista+ only
Private Type TOKEN_ELEVATION
    TokenIsElevated As Long
End Type

Private Type SHELLEXECUTEINFO
        cbSize        As Long
        fMask         As Long
        hwnd          As Long
        lpVerb        As String
        lpFile        As String
        lpParameters  As String
        lpDirectory   As String
        nShow         As Long
        hInstApp      As Long
        lpIDList      As Long     'Optional
        lpClass       As String   'Optional
        hkeyClass     As Long     'Optional
        dwHotKey      As Long     'Optional
        hIcon         As Long     'Optional
        hProcess      As Long     'Optional
End Type

Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_ELEVATION = 20
Private Const SW_SHOWNORMAL = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpSEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Option Explicit

Sub Main()
    
    Dim ocx As String
    
    'depandancy removed switched to listbox
    'this one is 1mb and probably already installed..very often used..
'    ocx = Environ("windir") & "\system32\MSCOMCTL.OCX"
'    If Not FileExists(ocx) Then
'        ocx = ocx = Environ("windir") & "\MSCOMCTL.OCX"
'        If Not FileExists(ocx) Then
'            MsgBox "Dependancy: MSCOMCTL.OCX not found on your system", vbExclamation
'            End
'        End If
'    End If
    
    'binary compat has been set on both of these libs
    'installs/registers if not already, alerts and dies if cant..
    
    If IsVistaPlus() Then
        If Not IsProcessElevated() Then
            If RunElevated(App.path & "\remoteLookup.exe", , Command) Then
                End
            Else
                WarnNonAdmin "Could not elevate, this should be run as admin"
                Exit Sub
            End If
        End If
    End If
    
    ensureOCXInstalled "MSWINSCK.OCX"  'small enough to just include
    ensureRegistered "proc_lib.clsCmnDlg", "procLib.dll"
    ensureRegistered "sppe.CPEEditor", "sppe.dll"
 
    Form1.Visible = True
    
End Sub

Function WarnNonAdmin(warning As String)
    On Error Resume Next
    Dim c
    
    Form1.List1.AddItem warning
    'For Each c In Form1.Controls
    '    c.Enabled = False
    'Next
    
    Form1.Visible = True
    
End Function

Public Function IsVistaPlus() As Boolean
    Dim osVersion As OSVERSIONINFO
    osVersion.dwOSVersionInfoSize = Len(osVersion)
    If GetVersionEx(osVersion) = 0 Then Exit Function
    If osVersion.dwPlatformId <> VER_PLATFORM_WIN32_NT Or osVersion.dwMajorVersion < 6 Then Exit Function
    IsVistaPlus = True
End Function

Function IsProcessElevated() As Boolean

    Dim fIsElevated As Boolean
    Dim dwError As Long
    Dim hToken As Long

    'Open the primary access token of the process with TOKEN_QUERY.
    If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hToken) = 0 Then GoTo cleanup
     
    Dim elevation As TOKEN_ELEVATION
    Dim dwSize As Long
    If GetTokenInformation(hToken, TOKEN_ELEVATION, elevation, Len(elevation), dwSize) = 0 Then
        'When the process is run on operating systems prior to Windows Vista, GetTokenInformation returns FALSE with the
        'ERROR_INVALID_PARAMETER error code because TokenElevation is not supported on those operating systems.
         dwError = Err.LastDllError
         GoTo cleanup
    End If

    fIsElevated = IIf(elevation.TokenIsElevated = 0, False, True)

cleanup:
    If hToken Then CloseHandle (hToken)
    'if ERROR_SUCCESS <> dwError then err.Raise
    IsProcessElevated = fIsElevated
End Function

Public Function RunElevated(ByVal FilePath As String, Optional ByVal hWndOwner As Long = 0, Optional EXEParameters As String = "") As Boolean
    Dim SEI As SHELLEXECUTEINFO
    
    On Error GoTo Err

    With SEI
        .cbSize = Len(SEI)
        .fMask = 0
        .lpFile = FilePath
        .nShow = SW_SHOWNORMAL
        .lpDirectory = GetParentFolder(FilePath)
        .lpParameters = EXEParameters
        .hwnd = hWndOwner
        .lpVerb = "runas"
    End With

    RunElevated = ShellExecuteEx(SEI)

    Exit Function
Err:
    RunElevated = False
End Function

Private Sub ensureOCXInstalled(fPath As String)
    On Error Resume Next
    Dim installPath As String
    
    If InStr(fPath, "\") < 1 Then fPath = App.path & "\" & fPath
    installPath = Environ("windir") & "\system32\" & FileNameFromPath(fPath)
    
    'we could also check version..but thats allot more code..
    If FileExists(installPath) Then Exit Sub
    
    If Not FileExists(fPath) Then GoTo errOut
      
    FileCopy fPath, installPath
    If Not FileExists(installPath) Then GoTo errOut
    Shell "regsvr32.exe """ & installPath & """", vbNormalFocus
    If FileExists(installPath) Then Exit Sub
     
errOut:
    MsgBox "Could not find OCX: " & FileNameFromPath(fPath)
    End
    
End Sub

Private Function ensureRegistered(progId As String, fPath As String) As Boolean

    On Error Resume Next
    
    Dim o As Object
    Dim installPath As String
    
    Set o = CreateObject(progId)
    
    If Not o Is Nothing Then
        ensureRegistered = True
        Exit Function
    End If
    
    If InStr(fPath, "\") < 1 Then fPath = App.path & "\" & fPath
    
    If Not FileExists(fPath) Then
       MsgBox "ActiveX Dll not found: " & fPath
       End
    Else
        installPath = Environ("windir") & "\system32\" & FileNameFromPath(fPath)
        FileCopy fPath, installPath
        Shell "regsvr32.exe """ & installPath & """", vbNormalFocus
        If Err.Number <> 0 Then
            MsgBox "failed to register OCX control? you may have to run with administrator privileges", vbInformation
        End If
    End If
    
    Set o = CreateObject(progId)
    
    If Not o Is Nothing Then
        ensureRegistered = True
        Exit Function
    End If
    
    MsgBox "Could not create ActiveX object: " & progId
    End
    
End Function




Function ReadFile(filename)
  Dim f, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function



Sub WriteFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub



Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function GetBaseName(path) As String
    Dim tmp, ub
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Private Function FileNameFromPath(fullPath) As String
    Dim tmp
    If InStr(fullPath, "\") > 0 Then
        tmp = Split(fullPath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    Else
        FileNameFromPath = fullPath
    End If
End Function

Function GetParentFolder(path) As String
    Dim tmp, ub
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function rpad(v, Optional l As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        rpad = v & String(l - x, char)
    Else
hell:
        rpad = v
    End If
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim i
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


