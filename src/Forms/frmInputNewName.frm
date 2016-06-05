VERSION 5.00
Begin VB.Form frmInputNewName 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input New Name"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSessionName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      TabIndex        =   0
      Text            =   "New Session"
      Top             =   825
      Width           =   4245
   End
   Begin VB.CommandButton bnkCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4380
      TabIndex        =   2
      Top             =   1365
      Width           =   1335
   End
   Begin VB.CommandButton btnOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   2775
      TabIndex        =   1
      Top             =   1365
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Session Name:"
      Height          =   330
      Left            =   255
      TabIndex        =   4
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The new Session will be created for you. Please enter the new Session name into textbox below:"
      Height          =   435
      Left            =   195
      TabIndex        =   3
      Top             =   165
      Width           =   5595
   End
End
Attribute VB_Name = "frmInputNewName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sSessionName As String
Private m_xResult As VbMsgBoxResult

Property Get SessionName( _
    ) As String
    
    SessionName = m_sSessionName
End Property

Property Let SessionName( _
    sNewName As String)
    
    m_sSessionName = sNewName
End Property

Property Get Result( _
    ) As VbMsgBoxResult
    
    Result = m_xResult
End Property

Private Sub bnkCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    m_sSessionName = txtSessionName.Text
    m_xResult = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub Form_Activate()
    SetFormColors
    txtSessionName.SelStart = 0
    txtSessionName.SelLength = Len(txtSessionName.Text)
End Sub

Private Sub Form_Load()
    m_sSessionName = "New Session"
    m_xResult = VbMsgBoxResult.vbCancel
End Sub

Private Sub SetFormColors()
    Me.BackColor = vbWindowBackground
    Me.ForeColor = vbWindowText
    Me.btnOK.BackColor = vbButtonFace
    Me.bnkCancel.BackColor = vbButtonFace
End Sub
