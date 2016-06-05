VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAmazon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' SYSTEMTIME structure for the Win32 API
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function HMACSHA1 Lib "BbHmacSha.dll" (ByVal sKey As String, ByVal sData As String) As Byte()
Private Declare Function gettimeofday Lib "BbHmacSha.dll" () As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Const C_sAws As String = "AWS "
Private Const C_sHttpGet As String = "GET"
Private Const C_sHttpPut As String = "PUT"
Private Const C_sHttpDelete As String = "DELETE"
Private Const C_sHttpVer As String = " HTTP/1.1"
Private Const C_sHttpContentType As String = "Content-Type"
Private Const C_sHttpAmazonDate As String = "x-amz-date"
Private Const C_sHttpAmazonMagic As String = "x-amz-magic"
Private Const C_sHttpAuthorization As String = "Authorization"
Private Const C_sHttpS3AmazonCom As String = "s3.amazonaws.com"

Private m_sAccessKeyId As String
Private m_sSecretAccessKey As String
Private m_sBucketName As String
Private m_sFileName As String
Private m_vaDays    ' day of the week names
Private m_vaMonths  ' month names
Private m_xOwnerForm As Form
Private m_nResult As Commands

Property Let AccessKeyId( _
    sAccessKeyId As String)
    m_sAccessKeyId = sAccessKeyId
End Property

Property Let SecretAccessKey( _
    sSecretAccessKey As String)
    m_sSecretAccessKey = sSecretAccessKey
End Property

Property Let BucketName( _
    sBucketName As String)
    m_sBucketName = sBucketName
End Property

Property Let FileName( _
    sFileName As String)
    m_sFileName = sFileName
End Property

Property Set OwnerForm( _
    xNewOwner As Form)
    Set m_xOwnerForm = xNewOwner
End Property

Property Get Result( _
    ) As Commands
    Result = m_nResult
End Property

Private Sub Class_Initialize()
    m_vaDays = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    m_vaMonths = Array("", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
End Sub

Public Function ListAllMyBuckets( _
    ) As String
    
    Dim sCanonical As String
    Dim sSignature As String
    Dim sAmazonDate As String

    sAmazonDate = aws_GetDatestamp()
    sCanonical = MakeCanonicalSring(C_sHttpGet, "", "", "", sAmazonDate, "", "/")
    sSignature = aws_GetSignature(m_sSecretAccessKey, sCanonical)
    ListAllMyBuckets = AmazonRequest(C_sHttpGet, sAmazonDate, sSignature, "/")
End Function

Public Function CreateBucket( _
    sBucketName As String) As String

    Dim sCanonical As String
    Dim sSignature As String
    Dim sAmazonDate As String
    Dim sResource As String
    
    sResource = "/" & sBucketName
    sAmazonDate = aws_GetDatestamp()
    sCanonical = MakeCanonicalSring(C_sHttpPut, "", "", "", sAmazonDate, "", sResource)
    sSignature = aws_GetSignature(m_sSecretAccessKey, sCanonical)
    CreateBucket = AmazonRequest(C_sHttpPut, sAmazonDate, sSignature, sResource)
End Function

Public Function DeleteBucket( _
    sBucketName As String) As String

    Dim sCanonical As String
    Dim sSignature As String
    Dim sAmazonDate As String
    Dim sResource As String
    
    sResource = "/" & sBucketName
    sAmazonDate = aws_GetDatestamp()
    sCanonical = MakeCanonicalSring(C_sHttpDelete, "", "", "", sAmazonDate, "", sResource)
    sSignature = aws_GetSignature(m_sSecretAccessKey, sCanonical)
    DeleteBucket = AmazonRequest(C_sHttpDelete, sAmazonDate, sSignature, sResource)
End Function

Public Function GetIesFile( _
    ) As String
    
    Dim sCanonical As String
    Dim sSignature As String
    Dim sAmazonDate As String
    Dim sResource As String

    sResource = "/" & m_sBucketName & "/" & m_sFileName
    sAmazonDate = aws_GetDatestamp()
    sCanonical = MakeCanonicalSring(C_sHttpGet, "", "", "", sAmazonDate, "", sResource)
    sSignature = aws_GetSignature(m_sSecretAccessKey, sCanonical)
    GetIesFile = AmazonRequest(C_sHttpGet, sAmazonDate, sSignature, sResource)
End Function

Public Function PutIesFile( _
    sIesFile As String) As String
    
    Dim sCanonical As String
    Dim sSignature As String
    Dim sAmazonDate As String
    Dim sResource As String

    sResource = "/" & m_sBucketName & "/" & m_sFileName
    sAmazonDate = aws_GetDatestamp()
    sCanonical = MakeCanonicalSring(C_sHttpPut, "", "text/plain", "", sAmazonDate, "", sResource)
    sSignature = aws_GetSignature(m_sSecretAccessKey, sCanonical)
    PutIesFile = AmazonRequest(C_sHttpPut, sAmazonDate, sSignature, sResource, "text/plain", sIesFile)
End Function

' Get a Datetime in GMT that can be used with the Amazon Web Services methods
' Returns a date in the format:   Thu, 17 Nov 2005 18:49:58 GMT
Private Function aws_GetDatestamp() As String
    Dim xSysTime As SYSTEMTIME

    GetSystemTime xSysTime ' get DateTime in GMT
    With xSysTime
        aws_GetDatestamp = m_vaDays(.wDayOfWeek) & ", " & Format(.wDay, "00") & _
        " " & m_vaMonths(.wMonth) & " " & Format(.wYear, "0000") & " " & _
        Format(.wHour, "00") & ":" & Format(.wMinute, "00") & ":" & _
        Format(.wSecond, "00") & " GMT"
    End With
End Function

Private Function MakeCanonicalSring( _
    sHttpVerb As String, _
    sContentMd5 As String, _
    sContentType As String, _
    sDate As String, _
    sAmazonDate As String, _
    sNxAmzMagic As String, _
    sResource As String) As String
    
    MakeCanonicalSring = sHttpVerb & vbLf & _
        sContentMd5 & vbLf & _
        sContentType & vbLf & _
        sDate & vbLf & _
        IIf(Len(sAmazonDate) = 0, "", C_sHttpAmazonDate & ":" & sAmazonDate & vbLf) & _
        IIf(Len(sNxAmzMagic) = 0, "", C_sHttpAmazonMagic & ":" & sNxAmzMagic & vbLf) & _
        sResource

End Function

'Build the signature used in Amazon Web Service method calls
Public Function aws_GetSignature( _
    ByVal sAccessKey As String, _
    ByVal sCanonical As String) As String
    Dim hDll As Long
    Dim baMac() As Byte
    Dim n As Integer
    Dim convString As String
    
    #If Trace = 1 Then
        TraceDebug "Try loading BbHmacSha.dll from '" & App.Path & "/BbHmacSha.dll'"
    #End If
    hDll = LoadLibrary(App.Path & "/BbHmacSha.dll") ' to load DLL from the current IESessions dir and not IE dir
    
    If 0 = hDll Then
        #If Trace = 1 Then
            TraceDebug "Can't load BbHmacSha.dll", , "E"
        #End If
        Exit Function
    Else
        #If Trace = 1 Then
            TraceDebug "Loaded BbHmacSha.dll"
        #End If
    End If
    baMac = HMACSHA1(sAccessKey, sCanonical)
    
    'convert HEX bytes to string
    n = 0
    For n = 0 To UBound(baMac)
        convString = convString & Chr$(baMac(n))
    Next n
    aws_GetSignature = Base64Encode(convString)
    
    FreeLibrary hDll
    Erase baMac
End Function

Private Function AmazonRequest( _
    sHttpVerb As String, _
    sAmazonDate As String, _
    sSignature As String, _
    sResource As String, _
    Optional sContentType As String = "", _
    Optional xData As Variant = vbEmpty) As String
    
    Dim pxProgress As frmProgress
    Set pxProgress = New frmProgress
    
    Dim xmlHTTP As New MSXML2.ServerXMLHTTP30
    With xmlHTTP
        .open bstrmethod:=sHttpVerb, bstrurl:="http://s3.amazonaws.com" & sResource, varAsync:=True
        If "" <> sContentType Then
            .setRequestHeader C_sHttpContentType, sContentType
        End If
        .setRequestHeader C_sHttpAmazonDate, sAmazonDate
        .setRequestHeader C_sHttpAuthorization, C_sAws & m_sAccessKeyId & ":" & sSignature
        .setRequestHeader "Host", C_sHttpS3AmazonCom
        Set pxProgress.Server = xmlHTTP
        pxProgress.Content = xData
        
        If GbIsIe7 Then
            ' <show modal patch for IE7>
            SetWindowPos pxProgress.hwnd, HWND_TOPMOST, 0, 0, 0, 0, G_nFlagsForTopmost
            pxProgress.Show vbModal, m_xOwnerForm
            ' </show modal patch for IE7>
        Else
            pxProgress.Show vbModal, m_xOwnerForm
        End If
        
        m_nResult = pxProgress.Command

        If Commands.Cancel = m_nResult Then
            AmazonRequest = "" ' user cancelled
        ElseIf Commands.Timeout = m_nResult Then
            AmazonRequest = "" ' operation timeout
        ElseIf Commands.OK = m_nResult Then
            AmazonRequest = .responseText ' OK
        Else
            AmazonRequest = "" ' unspecified error
        End If
        Set pxProgress.Server = Nothing
    End With
    
    Set pxProgress = Nothing
    Set xmlHTTP = Nothing
    
End Function

Private Function Base64Encode(inData)
  'RFC1521
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, i
  
  'For each group of 3 bytes
  For i = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, i, 1)) + _
      &H100 * MyASC(Mid(inData, i + 1, 1)) + MyASC(Mid(inData, i + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut & pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) & "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) & "="
  End Select
  Base64Encode = sOut
End Function

Private Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Set m_xOwnerForm = Nothing
End Sub
