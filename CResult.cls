VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Contributed by FireEye FLARE Team
'Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
'Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
'License: GPL
Option Explicit

Public errMsg As String
Public address As Long
Public dllName As String
Public exportName As String
Public ordinal As Long
Public hadError As Boolean
Public fullPath As String

Friend Sub setErr(msg)
    errMsg = msg
    hadError = True
End Sub

Function Dump() As String
    
    If hadError Then
        Dump = "Error: " & errMsg
    Else
        Dump = Hex(address) & " , " & exportName & " , " & ordinal & " , " & dllName
    End If
    
End Function
