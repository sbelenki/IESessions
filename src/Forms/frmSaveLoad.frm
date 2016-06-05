VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSaveLoad 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Save/Load Sessions"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmSessions 
      BackColor       =   &H8000000B&
      Caption         =   "Sessions:"
      Height          =   2925
      Left            =   135
      TabIndex        =   10
      Top             =   1635
      Width           =   6105
      Begin MSComctlLib.TreeView tvSessions 
         Height          =   2460
         Left            =   135
         TabIndex        =   11
         Top             =   315
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   4339
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         HotTracking     =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frmBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   60
      TabIndex        =   5
      Top             =   4665
      Width           =   6315
      Begin VB.CommandButton btnSaveCurrent 
         Appearance      =   0  'Flat
         Caption         =   "Save Current..."
         Default         =   -1  'True
         Height          =   330
         Left            =   60
         TabIndex        =   9
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton btnLoad 
         Appearance      =   0  'Flat
         Caption         =   "Load"
         Height          =   330
         Left            =   1650
         TabIndex        =   8
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton btnExit 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   330
         Left            =   4830
         TabIndex        =   7
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton btnDelete 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
         Height          =   330
         Left            =   3240
         TabIndex        =   6
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame frmTop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   6315
      Begin VB.Frame frmSeparator 
         Height          =   120
         Left            =   30
         TabIndex        =   2
         Top             =   885
         Width           =   6105
      End
      Begin VB.CommandButton btnImportExport 
         Appearance      =   0  'Flat
         Caption         =   "Import/Export..."
         Height          =   330
         Left            =   4815
         TabIndex        =   1
         Top             =   1185
         Width           =   1335
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   5700
         Top             =   -225
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSaveLoad.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSaveLoad.frx":015A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   390
         Picture         =   "frmSaveLoad.frx":02B4
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Label txtDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSaveLoad.frx":0ADD
         Height          =   630
         Left            =   1290
         TabIndex        =   4
         Top             =   105
         Width           =   4635
      End
      Begin VB.Label lblImpExp 
         BackStyle       =   0  'Transparent
         Caption         =   "Share your Sessions between computers and browser versions."
         Height          =   345
         Left            =   30
         TabIndex        =   3
         Top             =   1245
         Width           =   4545
      End
   End
End
Attribute VB_Name = "frmSaveLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Const TVM_GETIMAGELIST As Long = (&H1100 + 8)
Private Const TVM_SETIMAGELIST As Long = (&H1100 + 9)
Private Const TVSIL_NORMAL As Long = 0

Private m_xSessionsBag As PropsBag
Private m_saUrlsToLoad() As String
Private m_saUrlsToSave() As String
Private m_nCommands As Commands
Private m_sSections() As String
Private m_iSectionCount As Long
Private m_nCurrentSessionId As Long
Private m_nLastError As Integer

Private m_xCurrentNode As Node

Private Const C_nPadding As Long = 0
Private m_nFormMinWidth As Long
Private m_nFormMinHeight As Long

Property Get UrlsToLoad( _
    ) As String()
    
    UrlsToLoad = m_saUrlsToLoad
End Property

Property Let UrlsToSave( _
    ByRef saUrlsToSave() As String)
    
    m_saUrlsToSave = saUrlsToSave
End Property

Property Get Command( _
    ) As Commands
    
    Command = m_nCommands
End Property

Property Get LastError( _
    ) As Integer
    
    LastError = m_nLastError
End Property

Property Let CurrentSessionID( _
    nCurrentID As Long)
    
    m_nCurrentSessionId = nCurrentID
End Property

Private Sub btnDelete_Click()
    DeleteSelectedNode
End Sub

Private Sub btnExit_Click()
    m_nCommands = Commands.Cancel
    Unload Me
End Sub

Private Sub btnImportExport_Click()
    Dim pxImportExport As frmImportExport
    Set pxImportExport = New frmImportExport
    Set pxImportExport.CurrentSessionsTree = tvSessions
    Set pxImportExport.CurrentSessionsBag = m_xSessionsBag
    If GbIsIe7 Then
        ' <show modal patch for IE7>
        SetWindowPos pxImportExport.hwnd, HWND_TOPMOST, 0, 0, 0, 0, G_nFlagsForTopmost
        pxImportExport.Show vbModal, Me
        ' </show modal patch for IE7>
    Else
        pxImportExport.Show vbModal, Me
    End If
    
    If Commands.Load = pxImportExport.Command Then
        ' a new Session was imported - redraw Tree
        ShowSessions
        Set m_xCurrentNode = tvSessions.Nodes.Item(1)
        Set tvSessions.SelectedItem = m_xCurrentNode
    End If
    
    Set pxImportExport = Nothing
End Sub

Private Sub btnLoad_Click()
    LoadSelectedSession
End Sub

Private Sub btnSaveCurrent_Click()
    
    Dim pxNewName As frmInputNewName
    Set pxNewName = New frmInputNewName
    If GbIsIe7 Then
        ' <show modal patch for IE7>
        SetWindowPos pxNewName.hwnd, HWND_TOPMOST, 0, 0, 0, 0, G_nFlagsForTopmost
        pxNewName.Show vbModal, Me
        ' </show modal patch for IE7>
    Else
        pxNewName.Show vbModal, Me
    End If
    
    If VbMsgBoxResult.vbOK = pxNewName.Result Then
        SaveCurrentSession pxNewName.SessionName & _
        " (" & Format(Now, "MMM d, yyyy hh:mm:ss AMPM") & ")"
    End If
    
    Set pxNewName = Nothing
End Sub

Private Sub ShowSessions()
Dim sGet As String
Dim sKeys() As String
Dim iKeycount As Long
Dim iSection As Long
Dim iKey As Long
Dim lSect As Long
Dim nodxSessions As Node
Dim nodxSession As Node
Dim nodxKey As Node

    ClearTreeViewNodes tvSessions.hwnd
    
    tvSessions.LineStyle = tvwRootLines
    Set nodxSessions = tvSessions.Nodes.Add(, , , "Stored Sessions", 1, 1)
    nodxSessions.Expanded = True
    Set nodxSessions.Tag = New TreeNodeTag
    nodxSessions.Tag.NodeType = TreeNodeType.Root
    nodxSessions.Tag.Key = -1
    
    With m_xSessionsBag
        .EnumerateAllSections m_sSections(), m_iSectionCount
        For iSection = 1 To m_iSectionCount
            Set nodxSession = tvSessions.Nodes.Add(nodxSessions, tvwChild, , _
                m_sSections(iSection), 1, 1)
            Set nodxSession.Tag = New TreeNodeTag
            nodxSession.Tag.NodeType = TreeNodeType.Session
            nodxSession.Tag.Key = -1

            .Section = m_sSections(iSection)
            .EnumerateCurrentSection sKeys(), iKeycount
            For iKey = 1 To iKeycount
                .Key = sKeys(iKey)
                Set nodxKey = tvSessions.Nodes.Add(nodxSession, tvwChild, , .value, 2, 2)
                Set nodxKey.Tag = New TreeNodeTag
                nodxKey.Tag.NodeType = TreeNodeType.Uri
                nodxKey.Tag.Key = .Key
            Next iKey
        Next iSection
    End With
End Sub

Private Sub Form_Activate()
    SetFormColors
    ShowSessions
    Set m_xCurrentNode = tvSessions.Nodes.Item(1)
    Set tvSessions.SelectedItem = m_xCurrentNode
End Sub

Private Sub Form_Load()
    m_nFormMinWidth = Me.frmTop.Width * (Me.Width / Me.ScaleWidth)
    m_nFormMinHeight = (Me.frmTop.Height + Me.frmSessions.Height + _
        Me.frmBottom.Height + 2 * C_nPadding) * (Me.Height / Me.ScaleHeight)
    Me.Width = m_nFormMinWidth
    Me.Height = m_nFormMinHeight
    
    Set m_xSessionsBag = New PropsBag
    tvSessions.ImageList = imlIcons
    m_nCommands = Commands.None
    m_xSessionsBag.Path = App.Path & "\" & G_sSessionsBagFilename
End Sub

Private Sub Form_Resize()
    If Me.Width < m_nFormMinWidth Then Me.Width = m_nFormMinWidth
    If Me.Height < m_nFormMinHeight Then Me.Height = m_nFormMinHeight

    Dim nNewWidth As Single
    Dim nNewSessionsHeight As Single
    nNewWidth = Me.ScaleWidth - (2 * C_nPadding)
    nNewSessionsHeight = Me.ScaleHeight - _
        (2 * C_nPadding + Me.frmTop.Height + Me.frmBottom.Height)
    Me.frmTop.Move C_nPadding, C_nPadding, nNewWidth
    Me.frmSessions.Move C_nPadding, Me.frmTop.Height, nNewWidth, nNewSessionsHeight
    Me.frmBottom.Move C_nPadding, Me.ScaleHeight - _
        (C_nPadding + Me.frmBottom.Height), nNewWidth
    Me.tvSessions.Move 0, 0, Me.frmSessions.Width, Me.frmSessions.Height
    Me.frmSeparator.Move frmSeparator.Left, frmSeparator.Top, Me.frmTop.Width - (2 * C_nPadding)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_xSessionsBag = Nothing
End Sub

Private Sub tvSessions_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    If NewString <> "" Then
        If NewString <> m_xCurrentNode.Text Then
            Dim i As Long
            Dim nodx As Node
            
            If TreeNodeType.Session = m_xCurrentNode.Tag.NodeType Then
                With m_xSessionsBag
                    .Section = NewString
                    If (m_xCurrentNode.children > 0) And _
                        (TreeNodeType.Session = m_xCurrentNode.Tag.NodeType) Then
                        Set nodx = m_xCurrentNode.Child
                        For i = 1 To m_xCurrentNode.children
                            .Key = nodx.Tag.Key
                            .value = nodx.Text
                            Set nodx = nodx.Next
                        Next
                    End If
                    #If Trace = 1 Then
                        TraceDebug "Copied Section [" & m_xCurrentNode.Text & _
                            "] into Section [" & NewString & "]"
                    #End If
                    .Section = m_xCurrentNode.Text
                    .DeleteSection
                    #If Trace = 1 Then
                        TraceDebug "Deleted Section [" & .Section & "]"
                    #End If
                End With
            Else
                With m_xSessionsBag
                    .Section = m_xCurrentNode.Parent.Text
                    .Key = m_xCurrentNode.Tag.Key
                    .value = NewString
                    #If Trace = 1 Then
                        TraceDebug "Updated Section [" & .Section & "], Key '" & _
                            .Key & "', new Value '" & .value & "', old Value '" & _
                        m_xCurrentNode.Text
                    #End If
                End With
            End If ' If TreeNodeType.Session = m_xCurrentNode.Tag.NodeType Then
    
        End If ' If NewString <> m_xCurrentNode.Text Then
    Else
        Cancel = True
    End If ' If NewString <> "" Then
    
End Sub

Private Sub tvSessions_DblClick()
    LoadSelectedSession
End Sub

Private Sub tvSessions_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCodeConstants.vbKeyDelete = KeyCode Then
        DeleteSelectedNode
    End If
End Sub

Private Sub tvSessions_NodeClick(ByVal xNode As MSComctlLib.Node)
    Set m_xCurrentNode = xNode
End Sub

Private Function SaveCurrentSession( _
    sSessionName As String)
    Dim i As Integer
    Dim saUrlsToSave() As String
    
    saUrlsToSave = m_saUrlsToSave ' KB Q197190
    With m_xSessionsBag
        .Section = sSessionName
        For i = 1 To UBound(saUrlsToSave)
            .Key = Str(i)
            .value = m_saUrlsToSave(i)
        Next i
    End With
    ShowSessions ' refresh tree view
    
End Function

Private Function LoadSelectedSession()
    Dim i As Long
    Dim nodx As Node

    If (m_xCurrentNode.children > 0) And _
        (TreeNodeType.Session = m_xCurrentNode.Tag.NodeType) Then
        ' a session node selected
        ' load all URL from the session
        Set nodx = m_xCurrentNode.Child
        ReDim m_saUrlsToLoad(m_xCurrentNode.children)
        
        For i = 1 To m_xCurrentNode.children
            m_saUrlsToLoad(i) = nodx.Text
            'MsgBox nodx.Text
            Set nodx = nodx.Next
        Next
        m_nCommands = Commands.Load
        Unload Me
    ElseIf (TreeNodeType.Uri = m_xCurrentNode.Tag.NodeType) Then
        ' one URL selected, load only it
        ReDim m_saUrlsToLoad(1)
        m_saUrlsToLoad(1) = m_xCurrentNode.Text
        m_nCommands = Commands.Load
        Unload Me
    Else
        ' nothing selected or wrong selection - cancel
        m_nCommands = Commands.Cancel
    End If
    
End Function

Private Sub DeleteSelectedNode()

    Dim parentNode As Node
    Dim childNode As Node
    Dim i As Long
    
    If TreeNodeType.Root <> m_xCurrentNode.Tag.NodeType Then
        ' delete Session record
        If TreeNodeType.Session = m_xCurrentNode.Tag.NodeType Then
            With m_xSessionsBag
                .Section = m_xCurrentNode.Text
                .DeleteSection
            End With
        Else
            Dim sKey() As String
            Dim iCount As Long
            With m_xSessionsBag
                .Section = m_xCurrentNode.Parent.Text
                .EnumerateCurrentSection sKey(), iCount
                For i = 1 To iCount
                    .Key = sKey(i)
                    If m_xCurrentNode.Text = .value Then
                        .DeleteKey
                        Exit For
                    End If
                Next i
            End With
        End If
        ' delete selected node
        If TreeNodeType.Session = m_xCurrentNode.Tag.NodeType Then
            Set parentNode = m_xCurrentNode.Parent
            Set m_xCurrentNode.Tag = Nothing ' remove session tag
            
            ' remove children tags
            If (m_xCurrentNode.children > 0) Then
                Set childNode = m_xCurrentNode.Child
                For i = 1 To m_xCurrentNode.children
                    Set childNode.Tag = Nothing
                    Set childNode = childNode.Next
                Next
            End If
            
            tvSessions.Nodes.Remove (m_xCurrentNode.Index) ' delete current node
            Set m_xCurrentNode = parentNode
            tvSessions.Refresh
            Set tvSessions.SelectedItem = m_xCurrentNode
        Else
            Set parentNode = m_xCurrentNode.Parent
            Set m_xCurrentNode.Tag = Nothing ' remove URL's tag
            
            tvSessions.Nodes.Remove (m_xCurrentNode.Index) ' delete current node
            Set m_xCurrentNode = parentNode
            tvSessions.Refresh
            Set tvSessions.SelectedItem = m_xCurrentNode
        End If
    End If
End Sub

Private Sub SetFormColors()
    Me.BackColor = vbWindowBackground
    Me.ForeColor = vbWindowText
    Me.frmTop.BackColor = vbWindowBackground
    Me.frmBottom.BackColor = vbWindowBackground
    Me.frmSeparator.BackColor = vbWindowBackground
    Me.frmSessions.BackColor = vbWindowBackground
    Me.btnImportExport.BackColor = vbButtonFace
    Me.btnSaveCurrent.BackColor = vbButtonFace
    Me.btnLoad.BackColor = vbButtonFace
    Me.btnDelete.BackColor = vbButtonFace
    Me.btnExit.BackColor = vbButtonFace
    DoEvents
End Sub

