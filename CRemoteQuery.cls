VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRemoteQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Contributed by FireEye FLARE Team
'Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
'Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
'License: GPL
Option Explicit

Public WithEvents ws As Winsock
Attribute ws.VB_VarHelpID = -1
Public WithEvents timeout As Timer
Attribute timeout.VB_VarHelpID = -1

Private buffer As String
Private sentOk As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Event DataReceived(data As String, ByRef respondWith As String)

'it is kind of an annoying task to force synchronous behavior on top of
'an asnyc control..but it is what it is..I had a sync socket api class
'but no idea where it went :_(
'
'grep keyword for when I loose this latter: synchronous socket class

Function control(listen As Boolean) As Boolean
    On Error Resume Next
    If listen Then ws.listen Else ws.Close
    control = (Err.Description = 0)
End Function

Private Sub timeout_Timer()
    On Error Resume Next
    timeout.Enabled = False
    ws.Close
    ws.listen
End Sub

'remote client closed connection
Private Sub ws_Close()
    On Error Resume Next
    timeout.Enabled = False
    ws.Close
    ws.listen
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
    ws.Close
    timeout.Enabled = False
    timeout.Interval = 8000
    timeout.Enabled = True
    ws.Accept requestID
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim tmp As String, response As String
    
    On Error Resume Next
    
    ws.GetData tmp
    
    'we received the whole command in one event..
    If Right(tmp, 1) = vbCr Or Right(tmp, 1) = vbLf Then
        buffer = buffer & tmp
        tmp = vbCr
    End If
    
    If Left(tmp, 1) = vbCr Then
        buffer = Replace(Replace(buffer, vbCr, Empty), vbLf, Empty)
        If Len(buffer) = 0 Then Exit Sub
        RaiseEvent DataReceived(buffer, response)
        buffer = Empty
        If Len(response) > 0 Then
            ws.SendData response
            waitForSend
            Sleep 300
            DoEvents
        End If
        ws.Close
        ws.listen
    Else
        buffer = buffer & tmp
        'Debug.Print buffer
    End If
    
End Sub

Sub waitForSend()
    Dim startTime As Long
    
    startTime = GetTickCount
    sentOk = False
    Do While Not sentOk
        DoEvents
        If GetTickCount - startTime > 4000 Then Exit Do
    Loop
    
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print Description
End Sub

Private Sub ws_SendComplete()
    sentOk = True
End Sub
