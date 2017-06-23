VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRemoteExportClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'	Contributed by FireEye FLARE Team
'	Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
'   Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
'	License: GPL

Public remoteIP As String

Public WithEvents ws As Winsock
Attribute ws.VB_VarHelpID = -1
Private buffer As String
Private sentOk As Boolean
Private connectedOk As Boolean
Private responseOk As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const dbg As Boolean = True

Private Function connect(ip As String) As Boolean
    On Error Resume Next
    ws.Close
    Err.Clear
    ws.connect ip, 9000
    waitForConnect
    connect = (ws.State = sckConnected)
End Function

Private Sub ws_Close()
    responseOk = True
    If dbg Then Debug.Print "closed"
End Sub

Private Sub ws_Connect()
    connectedOk = True
    If dbg Then Debug.Print "connect"
End Sub

Function ScanProcess(pidOrName, response) As Boolean
    'On Error Resume Next
    buffer = Empty
    
    If Len(remoteIP) = 0 Then
        response = "Set Remote IP first"
        Exit Function
    End If
    
    If Not connect(remoteIP) Then
        response = "connection failed"
        Exit Function
    End If
    
    ws.SendData "attach:" & pidOrName & vbCr
    waitForSend
    
    If Not sentOk Then
        response = "send failed"
        Exit Function
    End If
    
    waitForResponse
    
    If Not responseOk Then
        response = "no response"
        Exit Function
    End If
    
    If Left(buffer, 3) = "ok:" Then
        ScanProcess = True
        buffer = Mid(buffer, 4)
    End If
    
    response = buffer
    
End Function

Function ResolveExport(apiOrAddress, response) As Boolean
    'On Error Resume Next
    buffer = Empty
    
    If Len(remoteIP) = 0 Then
        response = "Set Remote IP first"
        Exit Function
    End If
    
    If Not connect(remoteIP) Then
        response = "connection failed"
        Exit Function
    End If
    
    ws.SendData "resolve:" & apiOrAddress & vbCr
    waitForSend
    
    If Not sentOk Then
        response = "send failed"
        Exit Function
    End If
    
    waitForResponse
    
    If Not responseOk Then
        response = "no response"
        Exit Function
    End If
    
    If Left(buffer, 3) = "ok:" Then
        buffer = Mid(buffer, 4)
        If InStr(buffer, "Error:") < 1 Then
            ResolveExport = True
            buffer = Replace(buffer, " ", Empty)
        End If
    End If
    
    response = buffer
    
End Function

 

Private Sub ws_DataArrival(ByVal bytesTotal As Long)

    On Error Resume Next
    Dim tmp As String
    
    ws.GetData tmp
    buffer = buffer & tmp
    If dbg Then Debug.Print "data arrival"
    
End Sub

Sub waitForSend()
    Dim startTime As Long
    
    startTime = GetTickCount
    sentOk = False
    Do While Not sentOk
        DoEvents
        Sleep 5
        If GetTickCount - startTime > 4000 Then Exit Do
    Loop
    
End Sub

Sub waitForResponse()
    Dim startTime As Long
    
    startTime = GetTickCount
    responseOk = False
    Do While Not responseOk
        DoEvents
        Sleep 5
        If GetTickCount - startTime > 10000 Then Exit Do
    Loop
    
End Sub

Sub waitForConnect()
    Dim startTime As Long
    
    startTime = GetTickCount
    connectedOk = False
    Do While Not connectedOk
        DoEvents
        Sleep 5
        If GetTickCount - startTime > 4000 Then Exit Do
    Loop
    
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If dbg Then Debug.Print "Error: " & Description
End Sub

Private Sub ws_SendComplete()
    sentOk = True
    If dbg Then Debug.Print "sent"
End Sub

