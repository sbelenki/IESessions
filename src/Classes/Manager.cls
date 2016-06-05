VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Implements IObjectWithSiteTLB.IObjectWithSite
Implements IOleCommandTarget

Private WithEvents m_ie As InternetExplorer
Attribute m_ie.VB_VarHelpID = -1
Private m_Site As IUnknownVB
Private m_xMainFrameHwnd As Long
Private m_xShellDocObjectView As Long

Private Sub Class_Initialize()
    m_xMainFrameHwnd = 0
    #If Trace = 1 Then
        TraceDebug "Class_Initialize" & vbTab, m_xMainFrameHwnd
    #End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    #If Trace = 1 Then
        TraceDebug "Class_Terminate" & vbTab, m_xMainFrameHwnd
    #End If
    m_xMainFrameHwnd = 0
    Set m_ie = Nothing
    Set m_Site = Nothing
End Sub

Private Sub IObjectWithSite_GetSite( _
    ByVal priid As IObjectWithSiteTLB.GUIDPtr, ppvObj As IObjectWithSiteTLB.VOIDPtr)
    m_Site.QueryInterface priid, ppvObj
    #If Trace = 1 Then
        TraceDebug "IObjectWithSite_GetSite" & vbTab, m_xMainFrameHwnd
    #End If
End Sub

Private Sub IObjectWithSite_SetSite( _
    ByVal pSite As IObjectWithSiteTLB.IUnknownVB)
    
    If Not (pSite Is Nothing) Then
        Set m_Site = pSite ' save pointer for GetSite
        
        If GbIsAppCompatible Then
        Dim pServiceProvider As IServiceProvider
        Set pServiceProvider = m_Site
        
        Dim clsidWebApp As Guid
        Dim clsidWebBrowser2 As Guid
        
        'Query service provider to get IWebBrowser2 (InternetExplorer)
        CLSIDFromString StrPtr(IID_IWebBrowserApp), clsidWebApp
        CLSIDFromString StrPtr(IID_IWebBrowser2), clsidWebBrowser2
    
        'Set m_ie = pSite
        Set m_ie = pServiceProvider.QueryService(VarPtr(clsidWebApp), VarPtr(clsidWebBrowser2))
        m_xMainFrameHwnd = m_ie.hwnd
        #If Trace = 1 Then
            TraceDebug "IObjectWithSite_SetSite" & vbTab & "IESessions Loaded", m_xMainFrameHwnd
        #End If

            'GbIsAppCompatible = CheckApplicationCompatibility(m_ie)
        GbIsIe7 = CheckBrowserCompatibility(m_xMainFrameHwnd)
            #If Trace = 1 Then
                TraceDebug "GbIsIe7 set to " & CStr(GbIsIe7), m_xMainFrameHwnd
            #End If
        End If
        #If Trace = 1 Then
            TraceDebug "GbIsAppCompatible set to " & CStr(GbIsAppCompatible), m_xMainFrameHwnd
        #End If
    Else
        #If Trace = 1 Then
            TraceDebug "IObjectWithSite_SetSite" & vbTab & " IESessions UnLoaded", m_xMainFrameHwnd
        #End If
        Set m_Site = pSite
        Set m_ie = pSite
    End If

End Sub

Private Sub m_ie_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    ' The Proper place for URL recording!!!
    If GbIsAppCompatible Then
        #If Trace = 1 Then
            'TraceDebug "DocumentComplete" & vbTab & "URL: " & m_ie.Document.URL, m_xMainFrameHwnd
            TraceDebug "DocumentComplete" & vbTab & "URL: " & m_ie.LocationURL, m_xMainFrameHwnd
            'TraceDebug "DocumentComplete" & vbTab & "URL: " & CStr(URL), m_xMainFrameHwnd
        #End If
        
'        ' Find hosting browser window handler
'        FindActiveBrowserByFrame m_xMainFrameHwnd
        
        
'        If 0 = m_xShellDocObjectView Then
'            m_xShellDocObjectView = FindShellDocObjectViewByFrame(m_xMainFrameHwnd)
'        End If

    End If
End Sub

' Returns true if the hosting application is browsr
' Returns false if the hosting application is Windows Explorer
Public Function CheckApplicationCompatibility( _
        ByVal xIe As InternetExplorer _
    ) As Boolean
    If Not xIe Is Nothing Then
        #If Trace = 1 Then
            TraceDebug "Application Full Name: " & xIe.FullName, m_xMainFrameHwnd
        #End If
        If 0 <> InStr(1, xIe.FullName, "explorer.exe", vbTextCompare) Then
            CheckApplicationCompatibility = False ' a Windows Explorer shell is colling us
        Else
            CheckApplicationCompatibility = True
        End If
    Else
        CheckApplicationCompatibility = False
    End If
End Function

Public Function CheckBrowserCompatibility( _
    hWndFrame As Long) As Boolean

    Dim WndChild As Long

    CheckBrowserCompatibility = False

    If hWndFrame = 0 Then
        #If Trace = 1 Then
            TraceDebug "CheckBrowserCompatibility: Frame is 0!", hWndFrame, "E"
        #End If
    Else
        WndChild = FindWindowEx(hWndFrame, 0, "TabWindowClass", vbNullString)
        If WndChild <> 0 Then
            CheckBrowserCompatibility = True
        End If
    End If
End Function

Private Function WriteToLog( _
    sMsg As String, _
    Optional bOverwrite As Boolean = False)
    
    On Error GoTo ErrorHandler
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    If bOverwrite Then
        Open App.Path & "/IESessions.log" For Output As #iFileNumber
    Else
        Open App.Path & "/IESessions.log" For Append As #iFileNumber
    End If
    
    Print #iFileNumber, Str(m_xMainFrameHwnd); vbTab; Now; vbTab; sMsg

ErrorHandler:
    Close #iFileNumber
End Function

Private Sub IOleCommandTarget_QueryStatus( _
    ByVal pguidCmdGroup As VBShellLib.LPGUID, _
    ByVal cCmds As VBShellLib.ULONG, _
    ByVal prgCmds As VBShellLib.LPOLECMD, _
    ByVal pCmdText As VBShellLib.LPOLECMDTEXT)

    If GbIsAppCompatible Then
        #If Trace = 1 Then
            TraceDebug "IOleCommandTarget_QueryStatus", m_xMainFrameHwnd
        #End If
    
        Dim i As Integer
    
        For i = 0 To cCmds - 1
            Dim cmd As OLECMD
    
            CopyMemory cmd, ByVal prgCmds + (Len(cmd) * i), Len(cmd)
            cmd.cmdf = OLECMDF_SUPPORTED Or OLECMDF_ENABLED
            CopyMemory ByVal prgCmds + (Len(cmd) * i), cmd, Len(cmd)
        Next i
    End If
    
End Sub

Private Sub IOleCommandTarget_Exec( _
    ByVal pguidCmdGroup As VBShellLib.LPGUID, _
    ByVal nCmdID As VBShellLib.DWORD, _
    ByVal nCmdExecOpt As VBShellLib.DWORD, _
    pvaIn As Variant, _
    pvaOut As Variant)
    
    If GbIsAppCompatible Then
        Dim pxSessionsFrm As frmSaveLoad ' sessions form
        Dim nCommand As Commands ' returned command
        Dim saUrlsToLoad() As String
        Dim saUrlsToSave() As String
        Dim i As Integer
        
        #If Trace = 1 Then
            TraceDebug "IOleCommandTarget_Exec", m_xMainFrameHwnd
        #End If
    
        GetUrlsByFrame m_xMainFrameHwnd, saUrlsToSave
    
        Set pxSessionsFrm = New frmSaveLoad
        pxSessionsFrm.CurrentSessionID = m_xMainFrameHwnd
        pxSessionsFrm.UrlsToSave = saUrlsToSave
        If GbIsIe7 Then
            ' <show modal patch for IE7>
            EnableWindow m_xMainFrameHwnd, 0
            SetWindowPos pxSessionsFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, G_nFlagsForTopmost
            On Error Resume Next ' Not to stuck with disabled IE7
            pxSessionsFrm.Show vbModal
            Err.Clear
            On Error GoTo 0
            EnableWindow m_xMainFrameHwnd, 1
            ' </show modal patch for IE7>
        Else
            pxSessionsFrm.Show vbModal
        End If
        nCommand = pxSessionsFrm.Command
        If Commands.Load = nCommand Then
            saUrlsToLoad = pxSessionsFrm.UrlsToLoad
            For i = 1 To UBound(saUrlsToLoad)
                #If Trace = 1 Then
                    TraceDebug "Navigating to: " & saUrlsToLoad(i), m_xMainFrameHwnd
                #End If
                If GbIsIe7 Then
                    m_ie.Navigate2 saUrlsToLoad(i), navOpenInNewTab
                Else
                    m_ie.Navigate2 saUrlsToLoad(i), navOpenInNewWindow
                End If
            Next i
        End If
        Set pxSessionsFrm = Nothing
    End If

End Sub
