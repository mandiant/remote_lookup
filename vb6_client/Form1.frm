VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Remote Client Library Test"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtResponse 
      Height          =   330
      Left            =   1710
      TabIndex        =   9
      Top             =   1170
      Width           =   2670
   End
   Begin VB.CommandButton cmdResolve 
      Caption         =   "Resolve"
      Height          =   285
      Left            =   3420
      TabIndex        =   8
      Top             =   810
      Width           =   1005
   End
   Begin VB.TextBox txtResolve 
      Height          =   285
      Left            =   1710
      TabIndex        =   7
      Text            =   "getprocaddress"
      Top             =   810
      Width           =   1590
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   285
      Left            =   3420
      TabIndex        =   5
      Top             =   450
      Width           =   1005
   End
   Begin VB.TextBox txtPID 
      Height          =   285
      Left            =   1710
      TabIndex        =   4
      Text            =   "explorer"
      Top             =   450
      Width           =   1545
   End
   Begin VB.CommandButton cmdSetIP 
      Caption         =   "Set"
      Height          =   285
      Left            =   3420
      TabIndex        =   2
      Top             =   90
      Width           =   1005
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1710
      TabIndex        =   1
      Text            =   "192.168.0.10"
      Top             =   90
      Width           =   1545
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   90
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Response"
      Height          =   240
      Left            =   810
      TabIndex        =   10
      Top             =   1215
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "ResolveExport"
      Height          =   240
      Left            =   585
      TabIndex        =   6
      Top             =   855
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "PID or Process Name"
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   495
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP"
      Height          =   240
      Left            =   810
      TabIndex        =   0
      Top             =   135
      Width           =   870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'	Contributed by FireEye FLARE Team
'	Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
'   Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
'	License: GPL


Dim q As New CRemoteExportClient

Private Sub cmdResolve_Click()
    Dim ret As String
    q.ResolveExport txtResolve, ret
    txtResponse = ret
End Sub

Private Sub cmdScan_Click()
    Dim ret As String
    q.ScanProcess txtPID, ret
    txtResponse = ret
End Sub

Private Sub cmdSetIP_Click()
    q.remoteIP = txtIP
End Sub

Private Sub Form_Load()
    Set q.ws = ws
    cmdSetIP_Click
End Sub
