VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropsBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' =========================================================
' Class:    PropsBag
' =========================================================

Private m_sPath As String
Private m_sKey As String
Private m_sSection As String
Private m_sDefault As String
Private m_lLastReturnCode As Long

#If Win32 Then
    ' Profile String functions:
    Private Declare Function WritePrivateProfileString Lib "KERNEL32" _
        Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Long
        
    Private Declare Function GetPrivateProfileString Lib "KERNEL32" _
        Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As Any, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As Any, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#Else
    ' Profile String functions:
    Private Declare Function WritePrivateProfileString Lib "Kernel" _
        ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Integer
        
    Private Declare Function GetPrivateProfileString Lib "Kernel" _
        ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As Any, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Integer, _
        ByVal lpFileName As String) As Integer
#End If

Property Get LastReturnCode( _
    ) As Long
    
    LastReturnCode = m_lLastReturnCode
End Property

Property Get Success( _
    ) As Boolean
    
    Success = (m_lLastReturnCode <> 0)
End Property

Property Let Default( _
    sDefault As String)
    
    m_sDefault = sDefault
End Property

Property Get Default( _
    ) As String
    
    Default = m_sDefault
End Property

Property Let Path( _
    sPath As String)
    
    m_sPath = sPath
End Property

Property Get Path( _
    ) As String
    
    Path = m_sPath
End Property

Property Let Key( _
    sKey As String)
    
    m_sKey = sKey
End Property

Property Get Key( _
    ) As String
    
    Key = m_sKey
End Property

Property Let Section( _
    sSection As String)
    
    m_sSection = sSection
End Property

Property Get Section( _
    ) As String
    
    Section = m_sSection
End Property

Property Get Value( _
    ) As String
    
Dim sBuf As String
Dim iSize As String
Dim iRetCode As Integer

    sBuf = Space$(255)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(m_sSection, m_sKey, m_sDefault, sBuf, iSize, m_sPath)
    If (iSize > 0) Then
        Value = Left$(sBuf, iRetCode)
    Else
        Value = ""
    End If
End Property

Property Let Value( _
    sValue As String)
    
Dim iPos As Integer
    ' Strip chr$(0):
    iPos = InStr(sValue, Chr$(0))
    Do While iPos <> 0
        sValue = Left$(sValue, (iPos - 1)) & Mid$(sValue, (iPos + 1))
        iPos = InStr(sValue, Chr$(0))
    Loop
    m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, sValue, m_sPath)
End Property

Public Sub DeleteKey( _
    )
    m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, 0&, m_sPath)
End Sub

Public Sub DeleteSection( _
    )
    m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, 0&, m_sPath)
End Sub

Property Get INISection( _
    ) As String
    
Dim sBuf As String
Dim iSize As String
Dim iRetCode As Integer

    sBuf = Space$(8192)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(m_sSection, 0&, m_sDefault, sBuf, iSize, m_sPath)
    If (iSize > 0) Then
        INISection = Left$(sBuf, iRetCode)
    Else
        INISection = ""
    End If
End Property

Property Let INISection( _
    sSection As String)
    
    m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, sSection, m_sPath)
End Property

Property Get Sections( _
    ) As String
    
Dim sBuf As String
Dim iSize As String
Dim iRetCode As Integer

    sBuf = Space$(8192)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(0&, 0&, m_sDefault, sBuf, iSize, m_sPath)
    If (iSize > 0) Then
        Sections = Left$(sBuf, iRetCode)
    Else
        Sections = ""
    End If
End Property

Public Sub EnumerateCurrentSection( _
    ByRef sKey() As String, _
    ByRef iCount As Long)
    
Dim sSection As String
Dim iPos As Long
Dim iNextPos As Long
Dim sCur As String
    
    iCount = 0
    Erase sKey
    sSection = INISection
    If (Len(sSection) > 0) Then
        iPos = 1
        iNextPos = InStr(iPos, sSection, Chr$(0))
        Do While iNextPos <> 0
            sCur = Mid$(sSection, iPos, (iNextPos - iPos))
            If (sCur <> Chr$(0)) Then
                iCount = iCount + 1
                ReDim Preserve sKey(1 To iCount) As String
                sKey(iCount) = Mid$(sSection, iPos, (iNextPos - iPos))
                iPos = iNextPos + 1
                iNextPos = InStr(iPos, sSection, Chr$(0))
            End If
        Loop
    End If
End Sub

Public Sub EnumerateAllSections( _
    ByRef sSections() As String, _
    ByRef iCount As Long)
    
Dim sIniFile As String
Dim iPos As Long
Dim iNextPos As Long
Dim sCur As String
    
    iCount = 0
    Erase sSections
    sIniFile = Sections
    If (Len(sIniFile) > 0) Then
        iPos = 1
        iNextPos = InStr(iPos, sIniFile, Chr$(0))
        Do While iNextPos <> 0
            If (iNextPos <> iPos) Then
                sCur = Mid$(sIniFile, iPos, (iNextPos - iPos))
                iCount = iCount + 1
                ReDim Preserve sSections(1 To iCount) As String
                sSections(iCount) = sCur
            End If
            iPos = iNextPos + 1
            iNextPos = InStr(iPos, sIniFile, Chr$(0))
        Loop
    End If

End Sub

Public Sub SaveFormPosition( _
    ByRef frmThis As Object)
    
Dim sSaveKey As String
Dim sSaveDefault As String
On Error GoTo SaveError
    sSaveKey = Key
    If Not (frmThis.WindowState = vbMinimized) Then
        Key = "Maximised"
        Value = (frmThis.WindowState = vbMaximized) * -1
        If (frmThis.WindowState <> vbMaximized) Then
            Key = "Left"
            Value = frmThis.Left
            Key = "Top"
            Value = frmThis.Top
            Key = "Width"
            Value = frmThis.Width
            Key = "Height"
            Value = frmThis.Height
        End If
    End If
    Key = sSaveKey
    Exit Sub
SaveError:
    Key = sSaveKey
    m_lLastReturnCode = 0
    Exit Sub
End Sub

Public Sub LoadFormPosition( _
    ByRef frmThis As Object, _
    Optional ByRef lMinWidth = 3000, _
    Optional ByRef lMinHeight = 3000)
    
Dim sSaveKey As String
Dim sSaveDefault As String
Dim lLeft As Long
Dim lTOp As Long
Dim lWidth As Long
Dim lHeight As Long
On Error GoTo LoadError
    sSaveKey = Key
    sSaveDefault = Default
    Default = "FAIL"
    Key = "Left"
    lLeft = CLngDefault(Value, frmThis.Left)
    Key = "Top"
    lTOp = CLngDefault(Value, frmThis.Top)
    Key = "Width"
    lWidth = CLngDefault(Value, frmThis.Width)
    If (lWidth < lMinWidth) Then lWidth = lMinWidth
    Key = "Height"
    lHeight = CLngDefault(Value, frmThis.Height)
    If (lHeight < lMinHeight) Then lHeight = lMinHeight
    If (lLeft < 4 * Screen.TwipsPerPixelX) Then lLeft = 4 * Screen.TwipsPerPixelX
    If (lTOp < 4 * Screen.TwipsPerPixelY) Then lTOp = 4 * Screen.TwipsPerPixelY
    If (lLeft + lWidth > Screen.Width - 4 * Screen.TwipsPerPixelX) Then
        lLeft = Screen.Width - 4 * Screen.TwipsPerPixelX - lWidth
        If (lLeft < 4 * Screen.TwipsPerPixelX) Then lLeft = 4 * Screen.TwipsPerPixelX
        If (lLeft + lWidth > Screen.Width - 4 * Screen.TwipsPerPixelX) Then
            lWidth = Screen.Width - lLeft - 4 * Screen.TwipsPerPixelX
        End If
    End If
    If (lTOp + lHeight > Screen.Height - 4 * Screen.TwipsPerPixelY) Then
        lTOp = Screen.Height - 4 * Screen.TwipsPerPixelY - lHeight
        If (lTOp < 4 * Screen.TwipsPerPixelY) Then lTOp = 4 * Screen.TwipsPerPixelY
        If (lTOp + lHeight > Screen.Height - 4 * Screen.TwipsPerPixelY) Then
            lHeight = Screen.Height - lTOp - 4 * Screen.TwipsPerPixelY
        End If
    End If
    If (lWidth >= lMinWidth) And (lHeight >= lMinHeight) Then
        frmThis.Move lLeft, lTOp, lWidth, lHeight
    End If
    Key = "Maximised"
    If (CLngDefault(Value, 0) <> 0) Then
        frmThis.WindowState = vbMaximized
    End If
    Key = sSaveKey
    Default = sSaveDefault
    Exit Sub
LoadError:
    Key = sSaveKey
    Default = sSaveDefault
    m_lLastReturnCode = 0
    Exit Sub
End Sub

Public Function CLngDefault( _
    ByVal sString As String, _
    Optional ByVal lDefault As Long = 0) As Long
    
Dim lR As Long
On Error Resume Next
    lR = CLng(sString)
    If (Err.Number <> 0) Then
        CLngDefault = lDefault
    Else
        CLngDefault = lR
    End If
End Function