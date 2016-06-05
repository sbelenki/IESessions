VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportExport 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Import/Export Sessions"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameLocation 
      Caption         =   "Configure the place to store your browser Sessions:"
      Height          =   1650
      Left            =   255
      TabIndex        =   11
      Top             =   3120
      Width           =   4680
      Begin VB.CommandButton btnAmazonConfigure 
         Appearance      =   0  'Flat
         Caption         =   "Configure..."
         Height          =   330
         Left            =   3030
         TabIndex        =   15
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   60
         Left            =   270
         TabIndex        =   14
         Top             =   810
         Width           =   4095
      End
      Begin VB.OptionButton optLocation 
         Caption         =   "Local Computer/Network"
         Height          =   300
         Index           =   0
         Left            =   315
         TabIndex        =   13
         Top             =   360
         Width           =   2145
      End
      Begin VB.OptionButton optLocation 
         Caption         =   "Amazon S3 Service"
         Height          =   300
         Index           =   1
         Left            =   315
         TabIndex        =   12
         Top             =   1035
         Width           =   1740
      End
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   255
      TabIndex        =   9
      Top             =   2445
      Width           =   4680
   End
   Begin MSComctlLib.TreeView tvImportedSessions 
      Height          =   600
      Left            =   645
      TabIndex        =   8
      Top             =   5010
      Visible         =   0   'False
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   1058
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgSaveAsOpen 
      Left            =   270
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnClose 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   330
      Left            =   3600
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton btnExport 
      Appearance      =   0  'Flat
      Caption         =   "Export..."
      Height          =   330
      Left            =   3600
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame frameLine1 
      Height          =   120
      Left            =   255
      TabIndex        =   5
      Top             =   1065
      Width           =   4680
   End
   Begin VB.CommandButton btnImport 
      Appearance      =   0  'Flat
      Caption         =   "Import..."
      Height          =   330
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Import/Export Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   255
      TabIndex        =   10
      Top             =   2730
      Width           =   2025
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Export Sessions from Internet Explorer to use them on another computer."
      Height          =   480
      Left            =   255
      TabIndex        =   7
      Top             =   1770
      Width           =   3030
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Export Sessions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   255
      TabIndex        =   6
      Top             =   1335
      Width           =   1860
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Import Sessions previously exported from another computer or browser."
      Height          =   480
      Left            =   255
      TabIndex        =   4
      Top             =   450
      Width           =   3285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Import Sessions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   255
      TabIndex        =   3
      Top             =   120
      Width           =   1605
   End
End
Attribute VB_Name = "frmImportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sSelectedFilename As String
Private m_nCommands As Commands
Private m_xCurrentSessionsTree As TreeView
Private m_xCurrentSessionsBag As PropsBag

Private Enum EnLocations
    LocalNetwork
    Amazon
End Enum

Property Get Command( _
    ) As Commands
    
    Command = m_nCommands
End Property

Property Set CurrentSessionsTree( _
    ByRef tvTree As TreeView)
    Set m_xCurrentSessionsTree = tvTree
End Property

Property Set CurrentSessionsBag( _
    ByRef xSessionsBag As PropsBag)
    Set m_xCurrentSessionsBag = xSessionsBag
End Property

Private Sub btnAmazonConfigure_Click()
    Dim pxAmazonCfg As frmConfigureAmazon
    Set pxAmazonCfg = New frmConfigureAmazon
    If GbIsIe7 Then
        ' <show modal patch for IE7>
        SetWindowPos pxAmazonCfg.hwnd, HWND_TOPMOST, 0, 0, 0, 0, G_nFlagsForTopmost
        pxAmazonCfg.Show vbModal, Me
        ' </show modal patch for IE7>
    Else
        pxAmazonCfg.Show vbModal, Me
    End If
    
    If pxAmazonCfg.Command = Commands.OK Then
        ' Store Amazon Config
        StoreAmazonConfig pxAmazonCfg.AccessKeyId, pxAmazonCfg.SecretAccessKey, pxAmazonCfg.BucketName, pxAmazonCfg.FileName, G_nHardPassword
    ElseIf pxAmazonCfg.Command = Commands.Delete Then
        ' Delete Amazon Config
        DeleteAmazonConfig
    End If
    
    Set pxAmazonCfg = Nothing
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnImport_Click()
    If Me.optLocation(EnLocations.LocalNetwork).value Then
        ImportLocally
    ElseIf Me.optLocation(EnLocations.Amazon).value Then
        On Error Resume Next ' can be firewalled
        ImportFromAmazon
        If 0 <> Err Then
            Err.Clear
            MsgBox "Unspecified exception while importing from Amazon S3 Service."
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub btnExport_Click()
    If Me.optLocation(EnLocations.LocalNetwork).value Then
        ExportLocally
    ElseIf Me.optLocation(EnLocations.Amazon).value Then
        On Error Resume Next ' can be firewalled
        ExportToAmazon
        If 0 <> Err Then
            Err.Clear
            MsgBox "Error while exporting to Amazon S3 Service." & vbCrLf & _
                "Please configure your firewall to allow the connection."
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub Form_Activate()
    SetFormColors
End Sub

Private Sub Form_Load()
    m_nCommands = Commands.Cancel
    
    dlgSaveAsOpen.DefaultExt = "ies"
    dlgSaveAsOpen.Filter = "Session Files (*.ies)|*.ies"
    dlgSaveAsOpen.FilterIndex = 1
    dlgSaveAsOpen.Flags = cdlOFNOverwritePrompt
    dlgSaveAsOpen.InitDir = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\Desktop\"
    
    Me.optLocation(EnLocations.LocalNetwork).value = True
End Sub

Private Function ImportFromAmazon()
    Dim sAccessKeyId As String
    Dim sSecretAccessKey As String
    Dim sBucketName As String
    Dim sFileName As String
    Dim sIesFileAsString As String
    Dim sAmazonErrCode As String
    Dim sAmazonErrMsg As String
    
    If GetAmazonConfig(sAccessKeyId, sSecretAccessKey, sBucketName, sFileName, G_nHardPassword) Then
#If Trace = 1 Then
        TraceDebug "GetAmazonConfig:" & vbCrLf & _
            vbTab & "sAccessKeyId: " & sAccessKeyId & vbCrLf & _
            vbTab & "sSecretAccessKey: " & sSecretAccessKey & vbCrLf & _
            vbTab & "sBucketName: " & sBucketName & vbCrLf & _
            vbTab & "sFileName: " & sFileName
#End If
    Else
        MsgBox "Amazon S3 Service is not configured on this computer", vbCritical, "Import/Export Problem"
        Exit Function
    End If
    
    Dim pxAmazon As New clsAmazon
    pxAmazon.AccessKeyId = sAccessKeyId
    pxAmazon.SecretAccessKey = sSecretAccessKey
    pxAmazon.BucketName = sBucketName
    pxAmazon.FileName = sFileName
    Set pxAmazon.OwnerForm = Me
    sIesFileAsString = pxAmazon.GetIesFile()
    
    If Commands.OK = pxAmazon.Result Then
    If "" = sIesFileAsString Then
            m_nCommands = Commands.Cancel
            MsgBox "No IESessions data was found in the file imported from Amazon S3 Service", _
                vbCritical, "Import/Export Problem"
    Else
        If CheckForAmazonError(sIesFileAsString, sAmazonErrCode, sAmazonErrMsg) Then
                m_nCommands = Commands.Cancel
                MsgBox "Amazon S3 Service error: " & vbCrLf & _
                    "Error Code: " & sAmazonErrCode & vbCrLf & _
                    "Error Message: " & sAmazonErrMsg, _
                vbCritical, "Import/Export Error"
        Else
            If TryImportFromAmazon(sIesFileAsString) Then
                If MergeSessions Then
                    m_nCommands = Commands.Load
                    MsgBox "Successfully imported data from Amazon S3 Service."
                Else
                        m_nCommands = Commands.Cancel
                    MsgBox "Error while merging IESessions file from Amazon S3 Service.", _
                        vbCritical, "Import/Export Error"
                End If
            Else
                m_nCommands = Commands.Cancel
                    MsgBox "IESessions file from Amazon S3 Service is in the wrong format.", _
                    vbCritical, "Import/Export Error"
            End If
        End If
    End If
    ElseIf Commands.Timeout = pxAmazon.Result Then
        m_nCommands = Commands.Cancel
        MsgBox "Time out while importing from Amazon S3 Service." & vbCrLf & _
            "Please configure your firewall to allow the connection.", _
            vbCritical, "Import/Export Error"
    ElseIf Commands.None = pxAmazon.Result Then
        m_nCommands = Commands.Cancel
        MsgBox "Unspecified error while importing from Amazon S3 Service.", _
            vbCritical, "Import/Export Error"
    End If
    
    Set pxAmazon = Nothing
End Function

Private Function ExportToAmazon()
    Dim sAccessKeyId As String
    Dim sSecretAccessKey As String
    Dim sBucketName As String
    Dim sFileName As String
    Dim sIesFileAsString As String
    Dim sResult As String
    Dim sAmazonErrCode As String
    Dim sAmazonErrMsg As String
    
    If GetAmazonConfig(sAccessKeyId, sSecretAccessKey, sBucketName, sFileName, G_nHardPassword) Then
#If Trace = 1 Then
        TraceDebug "GetAmazonConfig:" & vbCrLf & _
            vbTab & "sAccessKeyId: " & sAccessKeyId & vbCrLf & _
            vbTab & "sSecretAccessKey: " & sSecretAccessKey & vbCrLf & _
            vbTab & "sBucketName: " & sBucketName & vbCrLf & _
            vbTab & "sFileName: " & sFileName
#End If
    Else
        MsgBox "Amazon S3 Service is not configured on this computer", vbCritical, "Import/Export Problem"
        Exit Function
    End If
    
    sIesFileAsString = GetSessionsBagAsString()
    
    Dim pxAmazon As New clsAmazon
    pxAmazon.AccessKeyId = sAccessKeyId
    pxAmazon.SecretAccessKey = sSecretAccessKey
    pxAmazon.BucketName = sBucketName
    pxAmazon.FileName = sFileName
    sResult = pxAmazon.PutIesFile(sIesFileAsString)
    
    If Commands.OK = pxAmazon.Result Then
    If "" = sResult Then
        m_nCommands = Commands.Save
        MsgBox "Successfully exported data to Amazon S3 Service."
    Else
        m_nCommands = Commands.Cancel
            If CheckForAmazonError(sIesFileAsString, sAmazonErrCode, sAmazonErrMsg) Then
                MsgBox "Amazon S3 Service error: " & vbCrLf & _
                    "Error Code: " & sAmazonErrCode & vbCrLf & _
                    "Error Message: " & sAmazonErrMsg, _
                    vbCritical, "Import/Export Error"
            Else
                MsgBox "Error while exporting data to Amazon S3 Service:" & vbCrLf & sResult, _
                    vbCritical, "Import/Export Error"
            End If
        End If
    ElseIf Commands.Timeout = pxAmazon.Result Then
        m_nCommands = Commands.Cancel
        MsgBox "Time out while exporting data to S3 Service." & vbCrLf & _
            "Please configure your firewall to allow the connection.", _
            vbCritical, "Import/Export Error"
    ElseIf Commands.None = pxAmazon.Result Then
        m_nCommands = Commands.Cancel
        MsgBox "Unspecified error while exporting data to Amazon S3 Service.", _
            vbCritical, "Import/Export Error"
    End If
    
    Set pxAmazon = Nothing
End Function

Private Function ImportLocally()
    dlgSaveAsOpen.FileName = ""
    dlgSaveAsOpen.DialogTitle = "Import From"
    dlgSaveAsOpen.ShowOpen
    m_sSelectedFilename = dlgSaveAsOpen.FileName
    If m_sSelectedFilename <> "" Then
        If TryImportFromLocal(m_sSelectedFilename) Then
            If MergeSessions Then
                m_nCommands = Commands.Load
                MsgBox "Successfully imported data from '" & m_sSelectedFilename & "'"
            Else
                MsgBox "Error while merging Sessions from '" & m_sSelectedFilename & "'"
            End If
        Else
            m_nCommands = Commands.Cancel
            MsgBox "Error while importing from '" & m_sSelectedFilename & "'"
        End If
    End If
End Function

Private Function ExportLocally()
    dlgSaveAsOpen.FileName = ""
    dlgSaveAsOpen.DialogTitle = "Export To"
    dlgSaveAsOpen.ShowSave
    m_sSelectedFilename = dlgSaveAsOpen.FileName
    If m_sSelectedFilename <> "" Then
        If TryExportIntoLocal(m_sSelectedFilename) Then
            m_nCommands = Commands.Save
            MsgBox "Successfully exported data into '" & m_sSelectedFilename & "'"
        Else
            m_nCommands = Commands.Cancel
            MsgBox "Error while exporting into '" & m_sSelectedFilename & "'"
        End If
    End If
End Function

Private Function TryExportIntoLocal( _
    sFileName As String) As Boolean
    
    Dim m_xNewSessionsBag As PropsBag
    
    TryExportIntoLocal = False
    
    On Error Resume Next
    
    Set m_xNewSessionsBag = New PropsBag
    m_xNewSessionsBag.Path = sFileName
    ' enumerate current sessions tree starting from Root
    TraverseSaveTree m_xCurrentSessionsTree.Nodes.Item(1), m_xNewSessionsBag, False
    
    If Err <> 0 Then
        Err.Clear
    Else
        TryExportIntoLocal = True ' no errors
    End If
    On Error GoTo 0
    
    Set m_xNewSessionsBag = Nothing ' close the file if opend

End Function

Private Function MergeSessions( _
    ) As Boolean
    
    MergeSessions = True
    
    On Error Resume Next
        ' enumerate current sessions tree starting from Root
        TraverseSaveTree tvImportedSessions.Nodes.Item(1), m_xCurrentSessionsBag, True
        If Err <> 0 Then
            MergeSessions = False
            Err.Clear
        End If
    On Error GoTo 0
End Function

Private Function TraverseSaveTree( _
    ByRef xNode As Node, _
    ByRef xSessionsBag As PropsBag, _
    bOverride As Boolean)
    
    Dim objSiblingNode As Node
    
    Set objSiblingNode = xNode
    Do
        If bOverride Then
            OverrideSessionNode objSiblingNode, xSessionsBag
        Else
            SaveSessionNode objSiblingNode, xSessionsBag
        End If
        If Not objSiblingNode.Child Is Nothing Then
            TraverseSaveTree objSiblingNode.Child, xSessionsBag, bOverride
        End If
        Set objSiblingNode = objSiblingNode.Next
    Loop While Not objSiblingNode Is Nothing

End Function

Private Function SaveSessionNode( _
    ByRef xNode As Node, _
    ByRef xSessionsBag As PropsBag)
    With xSessionsBag
        If TreeNodeType.Session = xNode.Tag.NodeType Then
            .Section = xNode.Text
        ElseIf TreeNodeType.Uri = xNode.Tag.NodeType Then
            .Key = xNode.Tag.Key
            .value = xNode.Text
        End If
    End With
End Function

Private Function OverrideSessionNode( _
    ByRef xNode As Node, _
    ByRef xSessionsBag As PropsBag)
    With xSessionsBag
        If TreeNodeType.Session = xNode.Tag.NodeType Then
            .Section = xNode.Text
            .DeleteSection
            .Section = xNode.Text
        ElseIf TreeNodeType.Uri = xNode.Tag.NodeType Then
            .Key = xNode.Tag.Key
            .value = xNode.Text
        End If
    End With
End Function

Private Function TryImportFromLocal( _
    sFileName As String) As Boolean
    
    Dim sGet As String
    Dim sKeys() As String
    Dim iKeycount As Long
    Dim iSection As Long
    Dim iKey As Long
    Dim lSect As Long
    Dim nodxSessions As Node
    Dim nodxSession As Node
    Dim nodxKey As Node
    Dim m_sSections() As String
    Dim m_iSectionCount As Long
    Dim m_xNewSessionsBag As PropsBag

    TryImportFromLocal = False

    On Error Resume Next
    
    Set m_xNewSessionsBag = New PropsBag
    m_xNewSessionsBag.Path = sFileName
    
    ClearTreeViewNodes tvImportedSessions.hwnd
    
    Set nodxSessions = tvImportedSessions.Nodes.Add(, , , "New Sessions")
    Set nodxSessions.Tag = New TreeNodeTag
    nodxSessions.Tag.NodeType = TreeNodeType.Root
    nodxSessions.Tag.Key = -1
    
    With m_xNewSessionsBag
        .EnumerateAllSections m_sSections(), m_iSectionCount
        For iSection = 1 To m_iSectionCount
            Set nodxSession = tvImportedSessions.Nodes.Add(nodxSessions, tvwChild, , _
                m_sSections(iSection))
            Set nodxSession.Tag = New TreeNodeTag
            nodxSession.Tag.NodeType = TreeNodeType.Session
            nodxSession.Tag.Key = -1

            .Section = m_sSections(iSection)
            .EnumerateCurrentSection sKeys(), iKeycount
            For iKey = 1 To iKeycount
                .Key = sKeys(iKey)
                Set nodxKey = tvImportedSessions.Nodes.Add(nodxSession, tvwChild, , .value)
                Set nodxKey.Tag = New TreeNodeTag
                nodxKey.Tag.NodeType = TreeNodeType.Uri
                nodxKey.Tag.Key = .Key
            Next iKey
        Next iSection
    End With
    
    Set m_xNewSessionsBag = Nothing ' close the file if opend
    
    Err.Clear ' clear error for a case
    On Error GoTo 0
    
    ' check if something was loaded
    If tvImportedSessions.Nodes.Count > 1 Then
        TryImportFromLocal = True
    End If
End Function

Private Function TryImportFromAmazon( _
    sData As String) As Boolean

    Dim OutFileNum As Integer
    Dim sTmpFilePath As String
    
    sTmpFilePath = App.Path & "\" & G_sTempImportFilename
    ' save IES data from Amazon in the temporary file
    OutFileNum = FreeFile
    Open sTmpFilePath For Binary Access Write Lock Write As OutFileNum
    Put OutFileNum, , sData
    Close OutFileNum
    
    TryImportFromAmazon = TryImportFromLocal(sTmpFilePath)
    
    Kill sTmpFilePath
End Function

Private Function GetSessionsBagAsString( _
    ) As String

    Dim nInFileNum As Integer
    Dim sInFileName As String
    Dim nFileSize As Long
    Dim baData() As Byte
    
    sInFileName = App.Path & "\" & G_sSessionsBagFilename
    nInFileNum = FreeFile
    nFileSize = FileLen(sInFileName)
    If nFileSize > 0 Then
        ReDim baData(nFileSize - 1)
    Else
        Exit Function
    End If
    Open sInFileName For Binary Access Read Shared As nInFileNum
    Get nInFileNum, , baData
    Close nInFileNum
    
    GetSessionsBagAsString = StrConv(baData, vbUnicode)
    Erase baData
End Function

Private Sub SetFormColors()
    Me.BackColor = vbWindowBackground
    Me.ForeColor = vbWindowText
    Me.frameLine1.BackColor = vbWindowBackground
    Me.Frame1.BackColor = vbWindowBackground
    Me.Frame2.BackColor = vbWindowBackground
    Me.frameLocation.BackColor = vbWindowBackground
    Me.optLocation(0).BackColor = vbWindowBackground
    Me.optLocation(1).BackColor = vbWindowBackground
    Me.btnImport.BackColor = vbButtonFace
    Me.btnExport.BackColor = vbButtonFace
    Me.btnAmazonConfigure.BackColor = vbButtonFace
    Me.btnClose.BackColor = vbButtonFace
End Sub
