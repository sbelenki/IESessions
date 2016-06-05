VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProgress 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1050
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   113
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   5000
   End
   Begin VB.CommandButton btnCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   330
      Left            =   1673
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_nTimerInterval As Long = 50&

Private m_nCommands As Commands
Private m_xmlHTTP As MSXML2.ServerXMLHTTP30
Private m_resolveTimeout As Long
Private m_connectTimeout As Long
Private m_operationTimeout As Long
Private m_sContent

Property Get Command( _
    ) As Commands
    
    Command = m_nCommands
End Property

Property Set Server( _
    ByRef pxNewServer As MSXML2.ServerXMLHTTP30)
    
    Set m_xmlHTTP = pxNewServer
End Property

Property Get Server( _
    ) As MSXML2.ServerXMLHTTP30
    
    Set Server = m_xmlHTTP
End Property

Property Let Content( _
    sNewContent)
    
    m_sContent = sNewContent
End Property

Private Sub btnCancel_Click()
    Timer1.Interval = 0
    m_xmlHTTP.abort
    m_nCommands = Commands.Cancel ' user canceled
    Unload Me
End Sub

Private Sub Form_Activate()
    SetFormColors
    m_xmlHTTP.setTimeouts m_resolveTimeout, m_connectTimeout, m_operationTimeout, m_operationTimeout
    m_xmlHTTP.send m_sContent
    Timer1.Interval = C_nTimerInterval
End Sub

Private Sub Form_Load()
    m_resolveTimeout = 30& * 1000&
    m_connectTimeout = 30& * 1000&
    m_operationTimeout = 15& * 1000&
    ProgressBar1.Max = (m_resolveTimeout / C_nTimerInterval) + _
        (m_connectTimeout / C_nTimerInterval) + _
        (m_operationTimeout / C_nTimerInterval)
        
    m_nCommands = Commands.None
    ProgressBar1.value = 0
End Sub

Private Sub Timer1_Timer()
    ProgressBar1.value = ProgressBar1.value + 1
    If ProgressBar1.value = ProgressBar1.Max Then
        Timer1.Interval = 0
        m_xmlHTTP.abort
        m_nCommands = Commands.Timeout ' operations timeout
        Unload Me
    End If
    If m_xmlHTTP.ReadyState = 4 Then
        Timer1.Interval = 0
        m_nCommands = Commands.OK
        Unload Me
    End If
End Sub

Private Sub SetFormColors()
    Me.BackColor = vbWindowBackground
    Me.ForeColor = vbWindowText
    Me.btnCancel.BackColor = vbButtonFace
End Sub
