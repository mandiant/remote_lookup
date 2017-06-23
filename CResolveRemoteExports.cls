VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResolveRemoteExports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Contributed by FireEye FLARE Team
'Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
'Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
'License: GPL
Option Explicit

Public proc As New CProcessLib
Public pe As New CPEEditor
Public dlls As Collection 'of CDll
Public total As Long
Public benchMark As String
Public errorMsg As String
Public abort As Boolean

Private startTime As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Enum eventType
    et_dllCount
    et_modCount
    et_nextMod
    et_nextDll
    et_zero
End Enum

Event Progress(et As eventType, value)

Property Get isLoaded() As Boolean
    If Not dlls Is Nothing Then isLoaded = True
End Property

Function ScanProcess(pid As Long) As Boolean
    
    Dim c As Collection
    Dim dll As CModule
    Dim exp As CExport
    Dim d As CDll
    Dim e As CDllExport
    
    On Error GoTo fail
    
    total = 0
    abort = False
    errorMsg = Empty
    StartBenchMark
    Set dlls = New Collection
    Set c = proc.GetProcessModules(pid)
    
    If c Is Nothing Then
        errorMsg = "Could not enumerate process modules insufficient permissions"
        Exit Function
    End If
    
    For Each dll In c
        Set d = New CDll
        d.base = dll.base
        d.path = dll.path
        dlls.Add d
    Next
    
    RaiseEvent Progress(et_dllCount, dlls.Count)
    
    For Each d In dlls

        If abort Then Err.Raise 1, "ScanProcess", "User Aborted Scan"
        
        If Not pe.LoadFile(d.path) Then
            d.peLoadFail = True
        Else
            If pe.exports.functions.Count > 0 Then
                RaiseEvent Progress(et_modCount, pe.exports.functions.Count)
                For Each exp In pe.exports.functions
                    If abort Then Err.Raise 1, "ScanProcess", "User Aborted Scan"
                    Set e = New CDllExport 'we make a full copy for simplicity..could be complications and mem hog issues if we just used the ones from pe lib..
                    e.ord = exp.FunctionOrdial
                    e.address = exp.FunctionAddress
                    e.name = exp.FunctionName
                    d.exports.Add e
                    RaiseEvent Progress(et_nextMod, 0)
                    total = total + 1
                Next
            End If
        End If
        
        RaiseEvent Progress(et_nextDll, 0)
        
    Next
    
    RaiseEvent Progress(et_zero, 0)
    benchMark = EndBenchMark
    ScanProcess = True
    
    Exit Function
fail:
    errorMsg = Err.Description
    RaiseEvent Progress(et_zero, 0)
    benchMark = EndBenchMark
    
End Function

 


'supports the following:
'    hexMemoryAddress,
'    GetProcAddress (case insensitive),
'    ws2_32@13,
'    ntdll!atoi or msvcrt.atoi
Function ResolveExport(ByVal find) As CResult
    
    Dim dllName As String
    Dim isOrd As Boolean
    Dim d As CDll
    Dim value
    Dim r As New CResult
    Dim tmp
    
    If dlls Is Nothing Then
        r.setErr "You must first scan a process"
        Set ResolveExport = r
        Exit Function
    End If
    
    If dlls.Count = 0 Then
        r.setErr "No dlls are loaded in this process?"
        Set ResolveExport = r
        Exit Function
    End If

    find = Trim(find)

    StartBenchMark
    If isValidHex(find, value) Then
        Set ResolveExport = internal_ResolveExport(value, True)
        benchMark = EndBenchMark
        RaiseEvent Progress(et_zero, 0)
        Exit Function
    End If
    
    find = Replace(find, "!", ".")
    value = Trim(find)
    
    If InStr(find, ".") > 0 Then
        tmp = Split(find, ".")
        dllName = Trim(tmp(0)) & ".dll"
        value = Trim(tmp(1))
    End If
    
    isOrd = (InStr(value, "@") > 0)
    
    If isOrd Then
        If Len(dllName) = 0 Then
            If Left(value, 1) = "@" Then
                r.setErr "Ordinal lookups require a dll name"
                Set ResolveExport = r
                Exit Function
            Else
                tmp = Split(value, "@")
                dllName = tmp(0) & ".dll"
                value = tmp(1)
            End If
        End If
        value = Replace(value, "@", Empty)
    End If
    
    Set ResolveExport = internal_ResolveExport(value, False, dllName, isOrd)
    benchMark = EndBenchMark
    RaiseEvent Progress(et_zero, 0)
    
End Function


Private Function internal_ResolveExport(ByVal find, isAddress As Boolean, Optional dllName As String, Optional isOrdinal As Boolean) As CResult
    
    On Error Resume Next
    Dim d As CDll
    Dim e As CDllExport
    Dim r As New CResult
    
    RaiseEvent Progress(et_dllCount, dlls.Count)
    
    For Each d In dlls
    
        If Len(dllName) > 0 Then
            If LCase(dllName) <> LCase(FileNameFromPath(d.path)) Then GoTo nextOne
        End If
        
        RaiseEvent Progress(et_modCount, pe.exports.functions.Count)
                
        For Each e In d.exports
            If isAddress Then
                If d.base + e.address = find Then
                    Set internal_ResolveExport = setReturn(d, e)
                    Exit Function
                End If
            Else
                If isOrdinal Then
                    If e.ord = find Then
                        Set internal_ResolveExport = setReturn(d, e)
                        Exit Function
                    End If
                Else
                    If LCase(e.name) = LCase(find) Then
                        Set internal_ResolveExport = setReturn(d, e)
                        Exit Function
                    End If
                End If
            End If
            RaiseEvent Progress(et_nextMod, 0)
        Next
        
nextOne:
        RaiseEvent Progress(et_nextDll, 0)
    Next
    
    If isAddress Then
        r.setErr "Address Lookup not found"
    ElseIf isOrdinal Then
        r.setErr "Ordinal " & find & " not found in " & dllName
    ElseIf Len(dllName) > 0 Then
        r.setErr "Export " & find & " not found in " & dllName
    Else
        r.setErr "Export " & find & " not found in any loaded dll. (case insensitive but spelling must match)"
    End If
    
    Set internal_ResolveExport = r
    
End Function
 
Private Function setReturn(d As CDll, e As CDllExport) As CResult

    Dim r As New CResult
    
    r.address = d.base + e.address
    r.dllName = FileNameFromPath(d.path)
    r.fullPath = d.path
    r.ordinal = e.ord
    r.exportName = e.name
    
    'If Len(e.name) > 0 Then
    '    r.exportName = getbasename(d.path) & "." &e.name
    'Else
    '    r.exportName = getbasename(d.path) & "@" & e.ord
    'End If
    
    Set setReturn = r
    RaiseEvent Progress(et_zero, 0)
    
End Function

Sub StartBenchMark()
    startTime = GetTickCount()
End Sub

Function EndBenchMark() As String
    Dim endTime As Long, loadTime As Long
    endTime = GetTickCount()
    loadTime = endTime - startTime
    EndBenchMark = loadTime / 1000 & " seconds"
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



Private Function GetBaseName(path) As String
    Dim ub As String, tmp
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Private Function isValidHex(ByVal test As String, ByRef outVal) As Boolean
    On Error Resume Next
    outVal = CLng("&h" & Trim(Replace(test, "0x", Empty)))
    isValidHex = (Err.Number = 0)
End Function

