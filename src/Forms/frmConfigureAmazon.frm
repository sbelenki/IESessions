VERSION 5.00
Begin VB.Form frmConfigureAmazon 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configure Amazon S3 Connection"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnDelete 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Default         =   -1  'True
      Height          =   330
      Left            =   180
      TabIndex        =   14
      Top             =   4305
      Width           =   1335
   End
   Begin VB.TextBox txtBucketName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1845
      TabIndex        =   11
      Text            =   "<Bucket Name>"
      Top             =   2115
      Width           =   4380
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1845
      TabIndex        =   10
      Text            =   "IESessions.ies"
      Top             =   3120
      Width           =   4380
   End
   Begin VB.CommandButton btnTestConnection 
      Appearance      =   0  'Flat
      Caption         =   "Test Connection"
      Height          =   330
      Left            =   1755
      TabIndex        =   8
      Top             =   4305
      Width           =   1335
   End
   Begin VB.TextBox txtSecretAccessKey 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1845
      TabIndex        =   6
      Text            =   "<Secret Access Key>"
      Top             =   1125
      Width           =   4380
   End
   Begin VB.CommandButton btnOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   330
      Left            =   3330
      TabIndex        =   2
      Top             =   4305
      Width           =   1335
   End
   Begin VB.CommandButton bnkCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4905
      TabIndex        =   1
      Top             =   4305
      Width           =   1335
   End
   Begin VB.TextBox txtAccessKeyId 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1845
      TabIndex        =   0
      Text            =   "<Access Key ID>"
      Top             =   630
      Width           =   4380
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Optionally, you can provide a file name that will be used as a key in the Amazon bucket."
      Height          =   450
      Left            =   195
      TabIndex        =   15
      Top             =   2640
      Width           =   6000
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Bucket Name:"
      Height          =   330
      Left            =   225
      TabIndex        =   13
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name (Optional):"
      Height          =   330
      Left            =   225
      TabIndex        =   12
      Top             =   3165
      Width           =   1605
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmConfigureAmazon.frx":0000
      Height          =   450
      Left            =   195
      TabIndex        =   9
      Top             =   1575
      Width           =   6000
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Secret Access Key:"
      Height          =   330
      Left            =   225
      TabIndex        =   7
      Top             =   1170
      Width           =   1440
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmConfigureAmazon.frx":00A8
      Height          =   480
      Left            =   195
      TabIndex        =   5
      Top             =   3675
      Width           =   6000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Please enter an Access Key ID and Secret Access Key for your Amazon S3 Service account into the textboxes below."
      Height          =   450
      Left            =   195
      TabIndex        =   4
      Top             =   150
      Width           =   6000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Access Key ID:"
      Height          =   330
      Left            =   225
      TabIndex        =   3
      Top             =   675
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfigureAmazon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sAccessKeyId As String
Private m_sSecretAccessKey As String
Private m_sBucketName As String
Private m_sFileName As String
Private m_nCommands As Commands

Property Get Command( _
    ) As Commands
    
    Command = m_nCommands
End Property

Property Get AccessKeyId( _
    ) As String
    AccessKeyId = m_sAccessKeyId
End Property

Property Let AccessKeyId( _
    sAccessKeyId As String)
    m_sAccessKeyId = sAccessKeyId
End Property

Property Get SecretAccessKey( _
    ) As String
    SecretAccessKey = m_sSecretAccessKey
End Property

Property Let SecretAccessKey( _
    sSecretAccessKey As String)
    m_sSecretAccessKey = sSecretAccessKey
End Property

Property Get BucketName( _
    ) As String
    BucketName = m_sBucketName
End Property

Property Let BucketName( _
    sBucketName As String)
    m_sBucketName = sBucketName
End Property

Property Get FileName( _
    ) As String
    FileName = m_sFileName
End Property

Property Let FileName( _
    sFileName As String)
    m_sFileName = sFileName
End Property

Private Sub bnkCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    m_nCommands = Commands.Delete
    Unload Me
End Sub

Private Sub btnOK_Click()
    m_nCommands = Commands.OK
    Unload Me
End Sub

Private Sub btnTestConnection_Click()
    ' TestConnection
    Dim sResult As String
    Dim sAmazonErrCode As String
    Dim sAmazonErrMsg As String
    Dim pxAmazon As New clsAmazon
    Dim i As Integer
    Dim isBucketExists As Boolean
   
    pxAmazon.AccessKeyId = Me.txtAccessKeyId.Text
    pxAmazon.SecretAccessKey = Me.txtSecretAccessKey.Text
    pxAmazon.BucketName = Me.txtBucketName.Text
    pxAmazon.FileName = G_sDefExportFilename
    Set pxAmazon.OwnerForm = Me
    sResult = pxAmazon.ListAllMyBuckets()

    If Commands.Cancel = pxAmazon.Result Then
        GoTo CleanAndExit ' user canceled
    ElseIf Commands.Timeout = pxAmazon.Result Then
        MsgBox "Time out while connection to Amazon S3 Service." & vbCrLf & _
            "Please configure your firewall to allow the connection.", _
            vbCritical, "Import/Export Error"
        GoTo CleanAndExit ' timeout
    ElseIf Commands.None = pxAmazon.Result Then
        MsgBox "Unspecified error while connection to Amazon S3 Service.", _
            vbCritical, "Import/Export Error"
        GoTo CleanAndExit ' unspecified error
    End If

#If Trace = 1 Then
        TraceDebug "TestConnection Result:" & vbCrLf & sResult
#End If

    If CheckForAmazonError(sResult, sAmazonErrCode, sAmazonErrMsg) Then
        If "InvalidArgument" = sAmazonErrCode Or _
            "InvalidAccessKeyId" = sAmazonErrCode Or _
            "SignatureDoesNotMatch" = sAmazonErrCode Then
            MsgBox "Amazon S3 Service error: " & vbCrLf & _
                "Wrong <Access Key ID> or <Secret Access Key>"
        Else
            MsgBox "Amazon S3 Service error: " & vbCrLf & _
                "Error Code: " & sAmazonErrCode & vbCrLf & _
                "Error Message: " & sAmazonErrMsg, _
                vbCritical, "Test Connection Error"
        End If
    Else
        Dim asListOfBuckets
        asListOfBuckets = ParseListOfBuckets(sResult)
        For i = 0 To UBound(asListOfBuckets)
            If Me.txtBucketName.Text = asListOfBuckets(i) Then
                isBucketExists = True
                Exit For
            End If
        Next i
        If Not isBucketExists Then
            Dim mbRes As VbMsgBoxResult
            mbRes = MsgBox("The bucket '" & Me.txtBucketName.Text & _
                "' does not exist at your account. Try to create one?", vbOKCancel, "Test Connection")
            If vbCancel = mbRes Then
                GoTo CleanAndExit ' nothing to test - user canceled
            End If

            sResult = pxAmazon.CreateBucket(Me.txtBucketName.Text)
            
            If CheckForAmazonError(sResult, sAmazonErrCode, sAmazonErrMsg) Then
                ' error in the case of the bucket with the name already exists
                MsgBox "Amazon S3 Service error: " & vbCrLf & _
                    "Error Code: " & sAmazonErrCode & vbCrLf & _
                    "Error Message: " & sAmazonErrMsg, _
                    vbCritical, "Test Connection Error"
            Else
                MsgBox "The new bucket was successfully created.", , "Test Connection"
            End If
        Else
            MsgBox "Connection to Amazon S3 Serice is OK.", , "Test Connection"
        End If ' If Not isBucketExists Then
    End If
    
CleanAndExit:
    Set pxAmazon = Nothing
End Sub

Private Function ParseListOfBuckets(sResult) As String()
    Dim nBucketsPos1 As Integer
    Dim nBucketsPos2 As Integer
    Dim sBucketsXml As String
    Dim sBucketNames As String
    Dim nCurPos1 As Integer
    Dim nCurPos2 As Integer
    
    nBucketsPos1 = InStr(1, sResult, "<Buckets>", vbTextCompare)
    If 0 = nBucketsPos1 Then Exit Function ' no <Buckets> element
    nBucketsPos1 = nBucketsPos1 + 9
    nBucketsPos2 = InStr(nBucketsPos1, sResult, "</Buckets>", vbTextCompare)
    sBucketsXml = Mid(sResult, nBucketsPos1, nBucketsPos2 - nBucketsPos1)
    nCurPos1 = 1
    Do Until nCurPos1 >= Len(sBucketsXml)
        nCurPos1 = InStr(nCurPos1, sBucketsXml, "<Name>", vbTextCompare)
        If 0 = nCurPos1 Then Exit Do ' no <Bucket> elements left
        nCurPos1 = nCurPos1 + 6
        nCurPos2 = InStr(nCurPos1, sBucketsXml, "</Name>", vbTextCompare)
        If "" <> sBucketNames Then sBucketNames = sBucketNames & "?" ' add separator
        sBucketNames = sBucketNames & Mid(sBucketsXml, nCurPos1, nCurPos2 - nCurPos1)
    Loop
    ' ? is an illegal in the buckets name, so it will serve here
    ' as a delimiter for Split function
    ParseListOfBuckets = Split(sBucketNames, "?", , vbTextCompare)
End Function

Private Sub Form_Activate()
    SetFormColors
    m_sAccessKeyId = Me.txtAccessKeyId.Text
    m_sSecretAccessKey = Me.txtSecretAccessKey.Text
    m_sBucketName = Me.txtBucketName.Text
    m_sFileName = G_sDefExportFilename
End Sub

Private Sub Form_Load()
    m_nCommands = Commands.Cancel
End Sub

Private Sub txtAccessKeyId_Change()
    m_sAccessKeyId = Trim(Me.txtAccessKeyId.Text)
End Sub

Private Sub txtAccessKeyId_LostFocus()
    Me.txtAccessKeyId.Text = Trim(Me.txtAccessKeyId.Text)
End Sub

Private Sub txtBucketName_Change()
    m_sBucketName = Trim(Me.txtBucketName.Text)
End Sub

Private Sub txtBucketName_LostFocus()
    Me.txtBucketName.Text = Trim(Me.txtBucketName.Text)
End Sub

Private Sub txtFileName_Change()
    m_sFileName = Trim(Me.txtFileName.Text)
End Sub

Private Sub txtFileName_LostFocus()
    Me.txtFileName.Text = Trim(Me.txtFileName.Text)
End Sub

Private Sub txtSecretAccessKey_Change()
    m_sSecretAccessKey = Trim(Me.txtSecretAccessKey.Text)
End Sub

Private Sub txtSecretAccessKey_LostFocus()
    Me.txtSecretAccessKey.Text = Trim(Me.txtSecretAccessKey.Text)
End Sub

Private Sub SetFormColors()
    Me.BackColor = vbWindowBackground
    Me.ForeColor = vbWindowText
    Me.btnDelete.BackColor = vbButtonFace
    Me.btnOK.BackColor = vbButtonFace
    Me.btnTestConnection.BackColor = vbButtonFace
End Sub
