VERSION 5.00
Begin VB.UserControl ucProgress 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   810
   ScaleWidth      =   4800
   Begin VB.Shape s 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   3  'Vertical Line
      Height          =   195
      Left            =   0
      Top             =   0
      Width           =   4155
   End
End
Attribute VB_Name = "ucProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright David Zimmer 2005
'site: http://sandsprite.com
'license GPL

'drop in replacement for mscomctl progressbar with some extras

Private m_max As Long
Private m_value As Long
Private lastRefresh As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub UserControl_Initialize()
      s.Visible = False
      s.Width = 0
      s.Height = UserControl.Height
End Sub

Private Sub UserControl_Resize()
   If Not Ambient.UserMode Then 'hosted on form in user IDE, not runtime
        s.Width = UserControl.Width
        s.Visible = True
   End If
End Sub

Sub reset()
    m_value = 0
    s.Width = 0
    s.Visible = False
End Sub

Property Get Max() As Long
    Max = m_max
End Property

Property Let Max(v As Long)
    If v < 0 Then v = 0
    Call reset
    m_max = v
End Property

Property Get Value() As Long
    Value = m_value
End Property

Property Let Value(v As Long)
    
    On Error Resume Next
    Dim maxWidth As Long
    Dim curWidth As Long
    Dim t As Long
    
    If v < 0 Then v = 0
    If v > m_max Then m_value = m_max Else m_value = v
    
    If v = 0 Then
        reset
        Exit Property
    End If
    
    If Not s.Visible Then s.Visible = True
    
    If v = m_max Then
        s.Width = UserControl.Width
    Else
        maxWidth = UserControl.Width
        curWidth = (m_value * maxWidth) / Max
        s.Width = curWidth
    End If
    
    t = GetTickCount
    If t - lastRefresh > 150 Then 'eliminate some flicker use less cpu in tight loops
        UserControl.Refresh
        DoEvents
    End If
    
    lastRefresh = t
    
End Property

Sub inc(Optional ticks As Long = 1)
    Value = Value + ticks
End Sub

Sub dec(Optional ticks As Long = 1)
    Value = Value - ticks
End Sub

Sub setPercent(precentage As Long)
    On Error Resume Next
    
    If precentage <= 0 Then
        Value = 0
        Exit Sub
    End If
        
    If precentage >= 100 Then
        Value = m_max
        Exit Sub
    End If
    
    Dim v As Long
    v = (percentage * m_max) / 100
    Value = v
    
End Sub

Property Get FillStyle() As Long
    FillStyle = s.FillStyle
End Property

Property Let FillStyle(v As Long)
    On Error Resume Next
    s.FillStyle = v
End Property

Property Get FillColor() As Long
    FillColor = s.FillColor
End Property

Property Let FillColor(v As Long)
    On Error Resume Next
    s.FillColor = v
End Property

Property Get BackColor() As Long
     BackColor = s.BackColor
End Property

Property Let BackColor(v As Long)
    On Error Resume Next
    s.BackColor = v
End Property

