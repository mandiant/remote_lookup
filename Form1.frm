VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "test UI for Resolve remote exports class..."
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucProgress pb2 
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   780
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   344
   End
   Begin Project1.ucProgress pb 
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   540
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   344
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   60
      TabIndex        =   11
      Top             =   1020
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Left            =   5265
      Top             =   4680
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   5760
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Allow Remote Queries (port 9000)"
      Height          =   285
      Left            =   90
      TabIndex        =   10
      Top             =   4725
      Width           =   2760
   End
   Begin VB.CommandButton cmdBulk 
      Caption         =   "Bulk"
      Height          =   375
      Left            =   5490
      TabIndex        =   8
      Top             =   3870
      Width           =   690
   End
   Begin VB.TextBox txtResult 
      Height          =   330
      Left            =   675
      TabIndex        =   4
      Top             =   4275
      Width           =   5505
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "Lookup"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   3870
      Width           =   780
   End
   Begin VB.TextBox txtLookup 
      Height          =   330
      Left            =   1890
      TabIndex        =   2
      Top             =   3870
      Width           =   2580
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select PID"
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1455
   End
   Begin VB.Label lblPID 
      Height          =   285
      Left            =   1665
      TabIndex        =   9
      Top             =   90
      Width           =   1140
   End
   Begin VB.Label lblBench 
      Height          =   285
      Left            =   3195
      TabIndex        =   7
      Top             =   90
      Width           =   2940
   End
   Begin VB.Label Label5 
      Caption         =   "Api name or hex address"
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
      Height          =   285
      Left            =   45
      TabIndex        =   6
      Top             =   3915
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   4320
      Width           =   510
   End
   Begin VB.Label Label2 
      Caption         =   "Dlls"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   495
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Contributed by FireEye FLARE Team
'Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
'Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
'License: GPL
Option Explicit

Dim WithEvents res As CResolveRemoteExports
Attribute res.VB_VarHelpID = -1
Dim dlg As New clsCmnDlg
Dim WithEvents remote As CRemoteQuery
Attribute remote.VB_VarHelpID = -1

Private p As CProcess

Private Sub Check1_Click()
    remote.control (Check1.value = 1)
End Sub

Private Sub cmdBulk_Click()
    Dim p As String
    Dim tmp() As String
    Dim r As CResult
    Dim fOut As String
    Dim okCnt As Long
    Dim msg As String
    Dim ret() As String
    Dim t
    
    On Error Resume Next
    
    If Not res.isLoaded Then
        txtResult = "Error: You must scan a process first.."
        Exit Sub
    End If
    
    p = dlg.OpenDialog(AllFiles)
    If Len(p) = 0 Then Exit Sub
    
    tmp = Split(ReadFile(p), vbCrLf)
    For Each t In tmp
        If Len(Trim(t)) > 0 Then
            Set r = res.ResolveExport(t)
            If Not r.hadError Then okCnt = okCnt + 1
            push ret, "ResolveExport(" & t & ") = " & r.Dump
        End If
    Next
    
    If AryIsEmpty(ret) Then
        msg = "No lookups were performed. format is one lookup value per line.."
    Else
        fOut = GetParentFolder(p) & "\" & GetBaseName(p) & "_results.txt"
        If FileExists(fOut) Then Kill fOut
        WriteFile fOut, Join(ret, vbCrLf)
        msg = "Complete " & okCnt & "/" & UBound(ret) + 1 & " found"
    End If
    
    lblBench = msg
    If Err.Number <> 0 Then txtResult = "Error: " & Err.Description
    If FileExists(fOut) Then Shell "notepad.exe """ & fOut & """", vbNormalFocus
    
End Sub

Private Sub Form_Load()
    Set res = New CResolveRemoteExports
    Set remote = New CRemoteQuery
    Set remote.ws = ws
    Set remote.timeout = Timer1
End Sub

Private Sub cmdLookup_Click()
    Dim r As CResult
    On Error Resume Next
    
    If Not res.isLoaded Then
        txtResult = "Error: You must scan a process first.."
        Exit Sub
    End If
    
    Set r = res.ResolveExport(txtLookup)
    lblBench = res.benchMark
    txtResult = r.Dump()
    
End Sub

Private Sub Command1_Click()
    
    If Command1.Caption = "Abort Scan" Then
        res.abort = True
        Exit Sub
    End If
    
    List1.Clear
    Set res = New CResolveRemoteExports
    Set p = res.proc.SelectProcess()
    If p Is Nothing Then Exit Sub
        
    If p.is64Bit Then
        List1.AddItem "Sorry currently only 32bit"
        Exit Sub
    End If
    
    Command1.Caption = "Abort Scan"
    DoScan p
    Command1.Caption = "Select PID"
    
End Sub

Sub DoScan(p As CProcess)

    Dim d As CDll
    Dim tmp As String
    
    List1.Clear
    
    lblPID = "pid: " & p.pid
    Me.Caption = "Examining " & p.path
    List1.AddItem "Scanning..."
    
    If Not res.ScanProcess(p.pid) Then
        List1.AddItem "Error: " & res.errorMsg
        Exit Sub
    End If
    
    If res.dlls Is Nothing Then
        List1.AddItem "No dlls found in process"
        Exit Sub
    End If
    
    List1.Clear
    List1.AddItem " Base    Exports  Path"
    
    For Each d In res.dlls
        tmp = Right("00000000" & Hex(d.base), 8) & "    "
        
        If d.peLoadFail Then
            tmp = tmp & rpad("Err", 6)
        Else
            tmp = tmp & rpad(d.exports.Count, 6)
        End If
        
        List1.AddItem tmp & d.path
        List1.Refresh
    Next
    
    lblBench = res.benchMark
    Me.Caption = res.dlls.Count & " dlls " & res.total & " export loaded"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    res.abort = True
    End
End Sub

Private Sub Label5_Click()
    
    MsgBox "ResolveExport supports the following:" & vbCrLf & _
            "    hex memory address," & vbCrLf & _
            "    case insensitive api name," & vbCrLf & _
            "    ws2_32@13," & vbCrLf & _
            "    ntdll!atoi or msvcrt.atoi" & vbCrLf & _
            vbCrLf & _
            "Bulk lookup loads a file one entry per line", vbInformation

End Sub


Private Sub remote_DataReceived(data As String, ByRef respondWith As String)
    
    Dim p2 As CProcess
    Dim c As Collection
    Dim cnt As Long
    Dim tmp
    
    tmp = Split(LCase(data), ":")
    
    If tmp(0) = "attach" Then
        If IsNumeric(tmp(1)) Then
            If Not res.proc.GetProcess(CLng(tmp(1)), p2) Then
                respondWith = "fail:GetProcess"
                Exit Sub
            Else
                If p2.is64Bit Then
                    respondWith = "fail:32bit Only"
                    Exit Sub
                End If
            End If
            'if we make it to here we have a valid process...
        Else
            If InStr(tmp(1), ".") < 1 Then tmp(1) = tmp(1) & ".exe"
            Set c = res.proc.GetRunningProcesses
            For Each p2 In c
                If LCase(p2.path) = tmp(1) Then cnt = cnt + 1
            Next
            If cnt = 0 Then
                respondWith = "fail:Process name not found"
                Exit Sub
            End If
            If cnt > 1 Then
                respondWith = "fail:Multiple Processes with name."
                Exit Sub
            End If
            For Each p2 In c
                If LCase(p2.path) = tmp(1) Then Exit For
            Next
            If p2.is64Bit Then
                respondWith = "fail:32bit Only"
                Exit Sub
            End If
            'if we make it to here we have a valid process...
        End If
        
        If Not p Is Nothing Then 'optimization for multiple script runs..
            If p.pid = p2.pid Then
                respondWith = "ok:already loaded"
                Exit Sub
            End If
        End If
        
        Set p = p2
        DoScan p
        respondWith = "ok:" & Me.Caption & " time:" & lblBench
        Exit Sub
    End If
    
    If tmp(0) = "resolve" Then
        txtLookup = tmp(1)
        cmdLookup_Click
        respondWith = "ok:" & txtResult
        Exit Sub
    End If
    
    

    respondWith = "fail:unknown command"
    Exit Sub
                    
                
End Sub

Private Sub res_Progress(et As eventType, value)
    
    On Error Resume Next
    
    If et = et_dllCount Then
        pb.Max = value
        pb.value = 0
    End If
    
    If et = et_modCount Then
        pb2.Max = value
        pb2.value = 0
    End If
    
    If et = et_zero Then
        pb2.value = 0
        pb.value = 0
    End If
    
    If et = et_nextDll Then pb.value = pb.value + 1
    If et = et_nextMod Then pb2.value = pb2.value + 1
    
    DoEvents
    
End Sub

Private Sub txtLookup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdLookup_Click
    End If
End Sub

