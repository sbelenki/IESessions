Attribute VB_Name = "basGlobals"
Option Explicit
Option Base 1

'----------------------------------------------------------
' UDTs
'----------------------------------------------------------
Private Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Public Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Enum Commands
    None
    Cancel
    OK
    Load
    Save
    Delete
    Timeout
End Enum

Private Type encFileHeader
    FLType1 As Byte ' = e
    FLType2 As Byte ' = n
    FLType3 As Byte ' = c
    Alg As Byte     ' always 2 for SHA1
    RndVal As Long  ' a random value (to enforce the password)
End Type

'----------------------------------------------------------
' Win32 Function Declarations
'----------------------------------------------------------
Private Declare Function GetClassName Lib "user32" _
   Alias "GetClassNameA" ( _
   ByVal hwnd As Long, _
   ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long

Private Declare Function EnumChildWindows Lib "user32" ( _
   ByVal hWndParent As Long, _
   ByVal lpEnumFunc As Long, _
   lParam As Long) As Long

Private Declare Function RegisterWindowMessage Lib "user32" _
   Alias "RegisterWindowMessageA" ( _
   ByVal lpString As String) As Long

Private Declare Function SendMessageTimeout Lib "user32" _
   Alias "SendMessageTimeoutA" ( _
   ByVal hwnd As Long, _
   ByVal Msg As Long, _
   ByVal wParam As Long, _
   lParam As Any, _
   ByVal fuFlags As Long, _
   ByVal uTimeout As Long, _
   lpdwResult As Long) As Long
   
Public Declare Function GetWindow Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal wCmd As Long) As Long

Private Declare Function ObjectFromLresult Lib "oleacc" ( _
   ByVal lResult As Long, _
   riid As UUID, _
   ByVal wParam As Long, _
   ppvObject As Any) As Long

Public Declare Function FindWindow Lib "user32" _
   Alias "FindWindowA" ( _
   ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long
   
Public Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Public Declare Function CLSIDFromString Lib "ole32.dll" _
    (ByVal lpszProgID As Long, pCLSID As Guid) As Long
    
Public Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
    pDest As Any, _
    pSource As Any, _
    ByVal ByteLen As Long)

#If Trace = 1 Then
    Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" ( _
        ByVal lpOutputString As String)
#End If

Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" _
        (ByVal hModule As Long, _
        ByVal lpFileName As String, _
        ByVal nSize As Long) As Long

Public Declare Function EnableWindow Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal fEnable As Long) As Long

Public Declare Function BringWindowToTop Lib "user32" ( _
    ByVal hwnd As Long) As Long

Public Declare Function SetWindowPos Lib "user32.dll" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cX As Long, _
    ByVal cY As Long, _
    ByVal wFlags As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

'----------------------------------------------------------
' Win32 constants
'----------------------------------------------------------
Private Const SMTO_ABORTIFHUNG = &H2
Private Const S_OK As Long = 0
Public Const IID_IWebBrowserApp = _
    "{0002DF05-0000-0000-C000-000000000046}"
Public Const IID_IWebBrowser2 = _
    "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"
Public Const IID_IShellBrowser = _
    "{000214e2-0000-0000-c000-000000000046}"
Public Const IID_IDockingWindowFrame = _
    "{47d2657a-7b27-11d0-8ca9-00a0c92dbfe8}"
Public Const navOpenInNewTab = &H800
Public Const navOpenInNewWindow = &H1
Public Const MAX_PATH = 260
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const HWND_TOPMOST  As Long = -1
Public Const HWND_NOTOPMOST  As Long = -2
Public Const SWP_NOMOVE  As Long = &H2
Public Const SWP_NOSIZE  As Long = &H1
' TreeView
Public Const WM_SETREDRAW As Long = &HB
Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Public Const TVGN_ROOT As Long = &H0
'----------------------------------------------------------
' program constants
'----------------------------------------------------------
Public Const G_sSessionsBagFilename As String = "SessionsBag.ies"
Public Const G_sDefExportFilename As String = "IESessions.ies"
Public Const G_sAmazonConfigFilename As String = "Amazon.cfg"
Public Const G_sTempImportFilename As String = "TmpImp.ies"
Public Const G_nFlagsForTopmost  As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const G_nHardPassword As String = "IESessions"

'----------------------------------------------------------
' program globals
'----------------------------------------------------------
Public GbIsAppCompatible As Boolean     ' is the HBO compatible with host (WebBrowser or Windows Explorer)
Public GbIsIe7 As Boolean               ' is the browser of version 7.x

'
' IEDOMFromhWnd
'
' Returns the IHTMLDocument interface from a WebBrowser window
'
' hWnd - Window handle of the control
'
Function GetIEFromHWND(ByVal hwnd As Long, ByRef xIe As Object) As IHTMLDocument
Dim IID_IHTMLDocument As UUID
Dim hWndChild As Long
Dim lRes As Long
Dim lMsg As Long
Dim hr As Long

   If hwnd <> 0 Then
      
      If Not IsIEServerWindow(hwnd) Then
      
         ' Find a child IE server window
         EnumChildWindows hwnd, AddressOf EnumChildProc, hwnd
         
      End If
      
      If hwnd <> 0 Then
            
         ' Register the message
         lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT")
            
         ' Get the object pointer
         Call SendMessageTimeout(hwnd, lMsg, 0, 0, _
                 SMTO_ABORTIFHUNG, 1000, lRes)

         If lRes Then
               
            ' Initialize the interface ID
            With IID_IHTMLDocument
               .Data1 = &H626FC520
               .Data2 = &HA41E
               .Data3 = &H11CF
               .Data4(0) = &HA7
               .Data4(1) = &H31
               .Data4(2) = &H0
               .Data4(3) = &HA0
               .Data4(4) = &HC9
               .Data4(5) = &H8
               .Data4(6) = &H26
               .Data4(7) = &H37
            End With
               
            ' Get the object from lRes
            hr = ObjectFromLresult(lRes, IID_IHTMLDocument, _
                     0, GetIEFromHWND)

            Dim clsidWebApp As Guid
            Dim clsidWebBrowser2 As Guid
            Dim clsidShellBrowser As Guid
            Dim clsidDockWndFrame As Guid

            CLSIDFromString StrPtr(IID_IWebBrowserApp), clsidWebApp
            CLSIDFromString StrPtr(IID_IWebBrowser2), clsidWebBrowser2
            CLSIDFromString StrPtr(IID_IShellBrowser), _
                 clsidShellBrowser
            CLSIDFromString StrPtr(IID_IDockingWindowFrame), _
                 clsidDockWndFrame

            If hr = S_OK Then ' S_OK result
                Dim pServiceProvider As IServiceProvider
                Set pServiceProvider = GetIEFromHWND.parentWindow
                Set xIe = pServiceProvider.QueryService( _
                    VarPtr(clsidWebApp), VarPtr(clsidWebBrowser2))
                If xIe Is Nothing Then
                    #If Trace = 1 Then
                        TraceDebug "xIE is not set!", hwnd, "E"
                    #End If
                End If
            End If
               
         End If

      End If
      
   End If

End Function

Private Function IsIEServerWindow(ByVal hwnd As Long) As Boolean
Dim lRes As Long
Dim sClassName As String

   ' Initialize the buffer
   sClassName = String$(100, 0)
   
   ' Get the window class name
   lRes = GetClassName(hwnd, sClassName, Len(sClassName))
   sClassName = Left$(sClassName, lRes)
   
   IsIEServerWindow = StrComp(sClassName, _
                      "Internet Explorer_Server", _
                      vbTextCompare) = 0
   
End Function

'
' Copy this function to a .bas module
'
Function EnumChildProc(ByVal hwnd As Long, lParam As Long) As Long
   
   If IsIEServerWindow(hwnd) Then
      lParam = hwnd
   Else
      EnumChildProc = 1
   End If
   
End Function

#If Trace = 1 Then ' Two Debug Functions

    Public Function GlobalWriteToLog( _
        sMsg As String, _
        Optional bOverwrite As Boolean = False)
        
        On Error GoTo ErrorHandler
        Dim iFileNumber As Integer
        iFileNumber = FreeFile
        If bOverwrite Then
            Open App.Path + "/IESessions.log" For Output As #iFileNumber
        Else
            Open App.Path + "/IESessions.log" For Append As #iFileNumber
        End If
    
        Print #iFileNumber, "GLOBAL LOG"; vbTab; Now; vbTab; sMsg
    
ErrorHandler:
        Close #iFileNumber
    End Function

    ' Debug Classes - second parameter:
    '   N - Normal
    '   W - Warning
    '   E - Error
    Public Function TraceDebug( _
        sMsg As String, _
        Optional nHwnd As Long = 0, _
        Optional sClass As String = "N")
        
        Dim sOutput As String
        Dim asStrArray
        Dim i As Integer
        
        asStrArray = Split(sMsg, vbCr + vbLf, , vbTextCompare)
        
        If UBound(asStrArray) = 1 Then
            sOutput = "IESessions" + vbTab + sClass + vbTab + _
                Str(nHwnd) + vbTab + CStr(Now) + vbTab + sMsg
            OutputDebugString sOutput
        Else
            sOutput = "IESessions" + vbTab + sClass + vbTab + _
                Str(nHwnd) + vbTab + CStr(Now) + vbTab + asStrArray(0)
            OutputDebugString sOutput
            For i = 1 To UBound(asStrArray)
                sOutput = "IESessions" + vbTab + sClass + vbTab + _
                    Str(nHwnd) + vbTab + CStr(Now) + vbTab + asStrArray(0)
                OutputDebugString sOutput
            Next i
        End If
    End Function

#End If

Public Function GetUrlsByFrame( _
    hWndFrame As Long, _
    ByRef saUrlsToSave() As String)
    
    Dim colUris As Collection
    Dim i As Integer
    
    If 0 <> hWndFrame Then
        Set colUris = New Collection
        If GbIsIe7 Then
            'GetUrlsByFrameIe7 hWndFrame, colUris
            GetAllIe7Urls colUris
        Else
            ' GetUrlsByFrameIe6 hWndFrame, colUris
            GetAllIe6Urls colUris
        End If
        
        If colUris.Count > 0 Then
            ReDim saUrlsToSave(colUris.Count)
            For i = 1 To colUris.Count
                saUrlsToSave(i) = colUris.Item(i)
            Next i
            While (colUris.Count)
                colUris.Remove (1)
            Wend
        End If
        Set colUris = Nothing
    End If
    
End Function

' Get all opend URL for all IE6 browsers
Public Function GetAllIe7Urls( _
    ByRef colUris As Collection)

    Dim hWndFrame As Long
    Dim hWndIe As Long
    Dim szClassName As String
    Dim lRes As Long
    Dim xIe As WebBrowser

    hWndFrame = FindWindow("IEFrame", vbNullString)

    While hWndFrame
        szClassName = String$(MAX_PATH, 0) ' init the buffer
        lRes = GetClassName(hWndFrame, szClassName, MAX_PATH)
        szClassName = Left$(szClassName, lRes)
        If 0 = StrComp(szClassName, "IEFrame", _
            vbTextCompare) Then
            GetUrlsByFrameIe7 hWndFrame, colUris
        End If ' if "TabWindowClass" class name
        hWndFrame = GetWindow(hWndFrame, GW_HWNDNEXT)
    Wend ' While hWndFrame
    
End Function

' Get all opend URL for all IE6 browsers
Public Function GetAllIe6Urls( _
    ByRef colUris As Collection)

    Dim hWndFrame As Long
    Dim hWndIe As Long
    Dim szClassName As String
    Dim lRes As Long
    Dim xIe As WebBrowser

    hWndFrame = FindWindow("IEFrame", vbNullString)

    While hWndFrame
        szClassName = String$(MAX_PATH, 0) ' init the buffer
        lRes = GetClassName(hWndFrame, szClassName, MAX_PATH)
        szClassName = Left$(szClassName, lRes)
        If 0 = StrComp(szClassName, "IEFrame", _
            vbTextCompare) Then
            hWndIe = GetWindow(hWndFrame, GW_CHILD)
            While hWndIe
                szClassName = String$(MAX_PATH, 0) ' reinit the buffer
                lRes = GetClassName(hWndIe, szClassName, MAX_PATH)
                szClassName = Left$(szClassName, lRes)
                If 0 = StrComp(szClassName, "Shell DocObject View", _
                    vbTextCompare) Then
                    hWndIe = FindWindowEx(hWndIe, 0, _
                        "Internet Explorer_Server", vbNullString)
                    If hWndIe <> 0 Then
                        Call GetIEFromHWND(hWndIe, xIe) 'Get IWebBrowser2 from Handle
                        If Not (xIe Is Nothing) Then
                            #If Trace = 1 Then
                                TraceDebug "Got IWebBrowser2 from Handle " + Str(hWndIe), hWndFrame, "W"
                            #End If
                            colUris.Add (xIe.LocationURL)
                        Else
                            #If Trace = 1 Then
                                TraceDebug "IE is Nothing!", hWndFrame, "E"
                            #End If
                        End If
                    Else
                        #If Trace = 1 Then
                            TraceDebug "No 'Internet Explorer_Server' found!", hWndFrame, "E"
                        #End If
                    End If
                End If
                hWndIe = GetWindow(hWndIe, GW_HWNDNEXT)
            Wend ' While hWndIe
        End If ' if "TabWindowClass" class name
        hWndFrame = GetWindow(hWndFrame, GW_HWNDNEXT)
    Wend ' While hWndFrame

End Function

Public Function GetUrlsByFrameIe7( _
    hWndFrame As Long, _
    ByRef colUris As Collection)
    
    Dim hWndTab As Long
    Dim hWndIe As Long
    Dim szClassName As String
    Dim lRes As Long
    Dim xIe As WebBrowser

    hWndTab = GetWindow(hWndFrame, GW_CHILD)

    While hWndTab
        szClassName = String$(MAX_PATH, 0) ' init the buffer
        lRes = GetClassName(hWndTab, szClassName, MAX_PATH)
        szClassName = Left$(szClassName, lRes)
        If 0 = StrComp(szClassName, "TabWindowClass", _
            vbTextCompare) Then
            hWndIe = GetWindow(hWndTab, GW_CHILD)
            While hWndIe
                szClassName = String$(MAX_PATH, 0) ' reinit the buffer
                lRes = GetClassName(hWndIe, szClassName, MAX_PATH)
                szClassName = Left$(szClassName, lRes)
                If 0 = StrComp(szClassName, "Shell DocObject View", _
                    vbTextCompare) Then
                    hWndIe = FindWindowEx(hWndIe, 0, _
                        "Internet Explorer_Server", vbNullString)
                    If hWndIe <> 0 Then
                        Call GetIEFromHWND(hWndIe, xIe) 'Get IWebBrowser2 from Handle
                        If Not (xIe Is Nothing) Then
                            #If Trace = 1 Then
                                TraceDebug "Got IWebBrowser2 from Handle " + Str(hWndIe), hWndFrame, "W"
                            #End If
                            colUris.Add (xIe.LocationURL)
                        Else
                            #If Trace = 1 Then
                                TraceDebug "IE is Nothing!", hWndFrame, "E"
                            #End If
                        End If
                    Else
                        #If Trace = 1 Then
                            TraceDebug "No 'Internet Explorer_Server' found!", hWndFrame, "E"
                        #End If
                    End If
                End If
                hWndIe = GetWindow(hWndIe, GW_HWNDNEXT)
            Wend ' While hWndIe
        End If ' if "TabWindowClass" class name
        hWndTab = GetWindow(hWndTab, GW_HWNDNEXT)
    Wend ' While hWndTab

End Function

Public Sub Main()
    Dim sModuleName As String
    sModuleName = String$(MAX_PATH, 0) ' Predefine string length
    GetModuleFileName 0, sModuleName, MAX_PATH
    If 0 <> InStr(1, sModuleName, "explorer.exe", vbTextCompare) Then
        GbIsAppCompatible = False
    Else
        GbIsAppCompatible = True
    End If
    App.HelpFile = App.Path + "\IESessions.chm"
End Sub

' Quicky clear the treeview identified by the hWnd parameter
Public Sub ClearTreeViewNodes(ByVal hwnd As Long)
    Dim hItem As Long
    
    ' lock the window update to avoid flickering
    SendMessageLong hwnd, WM_SETREDRAW, False, &O0

    ' clear the treeview
    Do
        hItem = SendMessageLong(hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
        If hItem <= 0 Then Exit Do
        SendMessageLong hwnd, TVM_DELETEITEM, &O0, hItem
    Loop
    
    ' unlock the window
    SendMessageLong hwnd, WM_SETREDRAW, True, &O0
End Sub

Public Function StoreAmazonConfig( _
    sAccessKeyId As String, _
    sSecretAccessKey As String, _
    sBucketName As String, _
    sFileName As String, _
    sLocalPassword As String)

    Dim InFileNum As Integer, OutFileNum As Integer, InBuff As String
    Dim OutBuff As String, PasswordCheck As String
    Dim FLHeader As encFileHeader, K As Long, Q As Long, HashLen As Long
    Dim sOutFileName As String
    Dim sOutString As String
    Dim nPosInStr As Long

    DeleteAmazonConfig
    
    OutFileNum = FreeFile
    sOutFileName = App.Path + "\" + G_sAmazonConfigFilename
    Open sOutFileName For Binary Access Write Lock Write As OutFileNum
    HashLen = 20
    Randomize
    FLHeader.FLType1 = Asc("e")
    FLHeader.FLType2 = Asc("n")
    FLHeader.FLType3 = Asc("c")
    FLHeader.Alg = CByte(2)
    FLHeader.RndVal = CLng(CLng(2 ^ 31 - 1) * Rnd) ' create a random value
    
    ' save file header
    Put OutFileNum, , FLHeader
    
    ' save hash of password with the random value (making it even more difficult to crack)
    ' also, when encrypting the same file 2 (or more) times, it will have different output
    ' every time (even if it's the same password)
    Put OutFileNum, , HASH(sLocalPassword & "?" & FLHeader.RndVal, False)

    ' create output string to store in the file
    sOutString = sAccessKeyId + vbLf + sSecretAccessKey + vbLf + _
        sBucketName + vbLf + sFileName + vbLf
    nPosInStr = 1
    K = Not FLHeader.RndVal ' make the random number a negative number
    Do Until nPosInStr >= Len(sOutString)
        If nPosInStr + HashLen > Len(sOutString) Then
            InBuff = String(Len(sOutString) - nPosInStr + 1, 0)
        Else
            InBuff = String(HashLen, 0)
        End If
        
        InBuff = Mid(sOutString, nPosInStr, Len(InBuff))
        ' get the hash string and XOR it with the input buffer, then save it into outputfile
        Put OutFileNum, , StrXOR(InBuff, Left(HASH(sLocalPassword & "?" & K, False), Len(InBuff)))
        
        K = K + 1 ' increment K for the next buffer
        ' move pointer in the output string
        nPosInStr = nPosInStr + HashLen
    Loop
    ' close output file
    Close OutFileNum
End Function

Public Function GetAmazonConfig( _
    sAccessKeyId As String, _
    sSecretAccessKey As String, _
    sBucketName As String, _
    sFileName As String, _
    sLocalPassword As String) As Boolean

    Dim InFileNum As Integer, OutFileNum As Integer, InBuff As String
    Dim OutBuff As String, PasswordCheck As String
    Dim FLHeader As encFileHeader, K As Long, Q As Long, HashLen As Long
    Dim InFileName As String
    Dim sInString As String
    Dim nPosInStr As Long
    
    GetAmazonConfig = True
    
    InFileNum = FreeFile
    InFileName = App.Path + "\" + G_sAmazonConfigFilename
    Open InFileName For Binary Access Read Lock Write As InFileNum
    
    Get InFileNum, , FLHeader
    If FLHeader.FLType1 = Asc("e") And FLHeader.FLType2 = Asc("n") And FLHeader.FLType3 = Asc("c") And FLHeader.Alg = 2 Then
        HashLen = 20
        PasswordCheck = String(HashLen, 0)
        Get InFileNum, , PasswordCheck ' get hashed password from the encrypted file
        ' check if current hashed password is the same as the hashed password in the file
        If PasswordCheck <> HASH(sLocalPassword & "?" & FLHeader.RndVal, False) Then
            'Password is incorrect
            GetAmazonConfig = False
            GoTo CloseAll
        End If
    Else
        GetAmazonConfig = False
        GoTo CloseAll
    End If
    
    K = Not FLHeader.RndVal ' make the random number a negative number
    Do Until Loc(InFileNum) >= LOF(InFileNum)
        If Loc(InFileNum) + HashLen > LOF(InFileNum) Then
            InBuff = String(LOF(InFileNum) - Loc(InFileNum), 0)
        Else
            InBuff = String(HashLen, 0)
        End If
        
        Get InFileNum, , InBuff

        ' get the hash string and XOR it with the input buffer, then save it into outputfile
        sInString = sInString + StrXOR(InBuff, Left(HASH(sLocalPassword & "?" & K, False), Len(InBuff)))
        
        K = K + 1 ' increment K for the next buffer
    Loop
    
    Dim asAmazonCfg
    asAmazonCfg = Split(sInString, vbLf, , vbTextCompare)
    If UBound(asAmazonCfg) < 3 Then ' could be 5 elems - with empty asAmazonCfg(4) elem
        GetAmazonConfig = False
        GoTo CloseAll
    Else
        sAccessKeyId = asAmazonCfg(0)
        sSecretAccessKey = asAmazonCfg(1)
        sBucketName = asAmazonCfg(2)
        sFileName = asAmazonCfg(3)
    End If

CloseAll:
    Close InFileNum

End Function
' Deletes Amazon Config file if exists
Public Function DeleteAmazonConfig()
    Dim sAmazonFile As String
    sAmazonFile = App.Path + "\" + G_sAmazonConfigFilename
    If Dir(sAmazonFile) <> "" Then
        Kill sAmazonFile
    End If
End Function

Private Function StrXOR(Str1 As String, Str2 As String) As String
    Dim K As Integer
    
    If Len(Str1) <> Len(Str2) Then Exit Function
    
    StrXOR = String(Len(Str1), 0)
    For K = 1 To Len(Str1)
        Mid$(StrXOR, K, 1) = Chr(Asc(Mid$(Str1, K, 1)) Xor Asc(Mid$(Str2, K, 1)))
    Next K
End Function

Private Function HASH(Str As String, Optional ByVal ReturnHex As Boolean = True) As String
    Dim Ret As String, K As Integer
    Dim cSHA As New clsSHA
    
    Ret = cSHA.SHA1(Str)
    Set cSHA = Nothing

    If ReturnHex Then ' return hashed string as hex
        HASH = Ret
    Else ' return hashed string as binary
        HASH = String(Len(Ret) \ 2, 0)
        
        For K = 1 To Len(HASH)
            Mid$(HASH, K, 1) = Chr(Val("&H" & Mid$(Ret, K * 2, 2)))
        Next K
    End If
End Function

Public Function CheckForAmazonError( _
    sAmazonResponse As String, _
    sAmazonErrorCode As String, _
    sAmazonErrorMessage As String) As Boolean
    
    Dim nErrorPos As Long
    Dim nCodePos1 As Long
    Dim nCodePos2 As Long
    Dim nMsgPos1 As Long
    Dim nMsgPos2 As Long
    
#If Trace = 1 Then
    TraceDebug "sAmazonResponse:" + vbCrLf + sAmazonResponse
#End If
    
    nErrorPos = InStr(1, sAmazonResponse, "<Error>", vbTextCompare)
    
    If 0 <> nErrorPos Then
        CheckForAmazonError = True
        nCodePos1 = InStr(1, sAmazonResponse, "<Code>", vbTextCompare) + 6
        nCodePos2 = InStr(nCodePos1, sAmazonResponse, "</Code>", vbTextCompare)
        sAmazonErrorCode = Mid(sAmazonResponse, nCodePos1, nCodePos2 - nCodePos1)
        nMsgPos1 = InStr(1, sAmazonResponse, "<Message>", vbTextCompare) + 9
        nMsgPos2 = InStr(nMsgPos1, sAmazonResponse, "</Message>", vbTextCompare)
        sAmazonErrorMessage = Mid(sAmazonResponse, nMsgPos1, nMsgPos2 - nMsgPos1)
    Else
        CheckForAmazonError = False
    End If
End Function

'--------------------------------------------------------------------------------

' <WARNING> The function doesn't work - gives error on about:tabs document
' Returns true if the Broser is IE7
' Returns false if the Browser is version less then IE7
''''Public Function CheckBrowserCompatibility_old( _
''''        ByVal xIe As WebBrowser _
''''    ) As Boolean
''''    Dim xHtmlDoc As HTMLDocument
''''    Dim sAppVersion As String
''''
''''    'On Error GoTo ErrorHandler
''''    CheckBrowserCompatibility_old = False
''''
''''    If Not xIe Is Nothing Then
''''
''''''''''''    '<TEST>
''''''''''''    Dim doc As IHTMLDocument
''''''''''''    Dim actIe As WebBrowser
''''''''''''    Dim activeIeWnd As Long
''''''''''''    activeIeWnd = FindActiveBrowserByFrame(m_xMainFrameHwnd)
''''''''''''    Set doc = GetIEFromHWND(activeIeWnd, actIe)
''''''''''''    ' gives <Access Denied> error on about:tabs page
''''''''''''    ' sAppVersion = doc.parentWindow.navigator.appVersion
''''''''''''    '</TEST>
''''
''''        If xIe.Type = "HTML Document" Then
''''            Set xHtmlDoc = xIe.Document
''''            If Not xHtmlDoc Is Nothing Then
''''                ' gives <Access Denied> error on about:tabs page
''''                ' sAppVersion = xHtmlDoc.parentWindow.navigator.appVersion
''''
''''                If sAppVersion <> "" Then
''''                    If 0 <> InStr(sAppVersion, "MSIE 7.") Then
''''                        CheckBrowserCompatibility_old = True
''''                    End If
''''                End If
''''            End If
''''
''''            TraceDebug "The hosting Browser Version is " + sAppVersion, m_xMainFrameHwnd
''''        Else
''''            TraceDebug "The hosting App is " + xIe.Type, m_xMainFrameHwnd
''''        End If
''''    End If
''''
''''    Exit Function
''''ErrorHandler:
''''    TraceDebug "CheckBrowserCompatibility_old Error" + xIe.Type, m_xMainFrameHwnd, "E"
''''End Function


'--------------------------------------------------------------------------------
' The function is OK, just not necessary

''''Public Function FindBrowserWindow( _
''''    ) As Long
''''    Dim Wnd As Long
''''    Dim WndChild As Long
''''    Dim IE As WebBrowser
''''
''''    Wnd = FindWindow("IEFrame", vbNullString)
''''    If Wnd = 0 Then
''''        #If Trace = 1 Then
''''            TraceDebug "TBD: Not a Browser - no IEFrame!", Wnd, "E"
''''        #End If
''''    Else
''''        If GbIsIe7 Then
''''            WndChild = FindWindowEx(Wnd, 0, "TabWindowClass", vbNullString)
''''            If WndChild <> 0 Then
''''                WndChild = FindWindowEx(WndChild, 0, "Shell DocObject View", vbNullString)
''''                If WndChild <> 0 Then
''''                    WndChild = FindWindowEx(WndChild, 0, "Internet Explorer_Server", vbNullString)
''''                    If WndChild <> 0 Then
''''                        #If Trace = 1 Then
''''                            TraceDebug "Found 'Internet Explorer_Server'!", Wnd, "W"
''''                        #End If
''''                        Call GetIEFromHWND(WndChild, IE) 'Get Iwebbrowser2 from Handle
''''                        If Not (IE Is Nothing) Then
''''                            #If Trace = 1 Then
''''                                TraceDebug "Got Iwebbrowser2 from Handle, WndChild:" + Str(WndChild), Wnd, "W"
''''                            #End If
''''                            FindBrowserWindow = IE.hwnd
''''                        Else
''''                            #If Trace = 1 Then
''''                                TraceDebug "IE is not set!", Wnd, "E"
''''                            #End If
''''                        End If
''''                    Else
''''                        #If Trace = 1 Then
''''                            TraceDebug "No 'Internet Explorer_Server'!", Wnd, "E"
''''                        #End If
''''                    End If
''''                Else
''''                    #If Trace = 1 Then
''''                        TraceDebug "No 'Shell DocObject View'!", Wnd, "E"
''''                    #End If
''''                End If
''''            Else
''''                #If Trace = 1 Then
''''                    TraceDebug "No 'TabWindowClass'!", Wnd, "E"
''''                #End If
''''            End If
''''        Else ' If GbIsIe7
''''            WndChild = FindWindowEx(Wnd, 0, "Shell DocObject View", vbNullString)
''''            If WndChild <> 0 Then
''''                WndChild = FindWindowEx(WndChild, 0, "Internet Explorer_Server", vbNullString)
''''                If WndChild <> 0 Then
''''                    #If Trace = 1 Then
''''                        TraceDebug "Found 'Internet Explorer_Server'!", Wnd, "W"
''''                    #End If
''''                    Call GetIEFromHWND(WndChild, IE) 'Get Iwebbrowser2 from Handle
''''                    If Not (IE Is Nothing) Then
''''                        #If Trace = 1 Then
''''                            TraceDebug "Got Iwebbrowser2 from Handle, WndChild: " + Str(WndChild), Wnd, "W"
''''                        #End If
''''                    Else
''''                        #If Trace = 1 Then
''''                            TraceDebug "IE is not set!", Wnd, "E"
''''                        #End If
''''                    End If
''''                Else
''''                    #If Trace = 1 Then
''''                        TraceDebug "No 'Internet Explorer_Server'!", Wnd, "E"
''''                    #End If
''''                End If
''''            Else
''''                #If Trace = 1 Then
''''                    TraceDebug "No 'Shell DocObject View'!", Wnd, "E"
''''                #End If
''''            End If
''''        End If ' If GbIsIe7
''''    End If
''''End Function

'--------------------------------------------------------------------------------
' The function is OK, just not necessary

''''Public Function GetUrlsByFrameIe6( _
''''    hWndFrame As Long, _
''''    ByRef colUris As Collection)
''''
''''    Dim hWndShell As Long
''''    Dim hWndIe As Long
''''    Dim szClassName As String
''''    Dim lRes As Long
''''    Dim xIe As WebBrowser
''''
''''    hWndShell = GetWindow(hWndFrame, GW_CHILD)
''''    While hWndShell
''''        szClassName = String$(MAX_PATH, 0) ' reinit the buffer
''''        lRes = GetClassName(hWndShell, szClassName, MAX_PATH)
''''        szClassName = Left$(szClassName, lRes)
''''        If 0 = StrComp(szClassName, "Shell DocObject View", _
''''            vbTextCompare) Then
''''            hWndIe = FindWindowEx(hWndShell, 0, _
''''                "Internet Explorer_Server", vbNullString)
''''            If hWndIe <> 0 Then
''''                Call GetIEFromHWND(hWndIe, xIe) 'Get IWebBrowser2 from Handle
''''                If Not (xIe Is Nothing) Then
''''                    #If Trace = 1 Then
''''                        TraceDebug "Got IWebBrowser2 from Handle " + Str(hWndIe), hWndFrame, "W"
''''                    #End If
''''                    colUris.Add (xIe.LocationURL)
''''                Else
''''                    #If Trace = 1 Then
''''                        TraceDebug "IE is Nothing!", hWndFrame, "E"
''''                    #End If
''''                End If
''''            Else
''''                #If Trace = 1 Then
''''                    TraceDebug "No 'Internet Explorer_Server' found!", hWndFrame, "E"
''''                #End If
''''            End If
''''        End If
''''        hWndShell = GetWindow(hWndShell, GW_HWNDNEXT)
''''    Wend ' While hWndShell
''''
''''End Function

'--------------------------------------------------------------------------------
' The function is OK, just not necessary

''''Public Function FindShellDocObjectViewByFrame( _
''''    hWndFrame As Long _
''''    ) As Long
''''
''''    Dim WndChild As Long
''''
''''    If hWndFrame = 0 Then
''''        #If Trace = 1 Then
''''            TraceDebug "FindActiveBrowserFromFrame: Frame is 0", hWndFrame, "E"
''''        #End If
''''        FindShellDocObjectViewByFrame = 0
''''    Else
''''        If GbIsIe7 Then
''''            WndChild = FindWindowEx(hWndFrame, 0, "TabWindowClass", vbNullString)
''''            If WndChild <> 0 Then
''''                WndChild = FindWindowEx(WndChild, 0, "Shell DocObject View", vbNullString)
''''                If WndChild <> 0 Then
''''                    FindShellDocObjectViewByFrame = WndChild
''''                    #If Trace = 1 Then
''''                        TraceDebug "Got Shell DocObject View by Frame, WndChild:" + Str(WndChild), hWndFrame, "W"
''''                    #End If
''''                Else
''''                    FindShellDocObjectViewByFrame = 0
''''                    #If Trace = 1 Then
''''                        TraceDebug "No 'Shell DocObject View'!", hWndFrame, "E"
''''                    #End If
''''                End If
''''            Else
''''                #If Trace = 1 Then
''''                    TraceDebug "No 'TabWindowClass'!", hWndFrame, "E"
''''                #End If
''''            End If
''''        Else ' If GbIsIe7
''''            WndChild = FindWindowEx(hWndFrame, 0, "Shell DocObject View", vbNullString)
''''            If WndChild <> 0 Then
''''                FindShellDocObjectViewByFrame = WndChild
''''                #If Trace = 1 Then
''''                    TraceDebug "Got Shell DocObject View by Frame, WndChild:" + Str(WndChild), hWndFrame, "W"
''''                #End If
''''            Else
''''                FindShellDocObjectViewByFrame = 0
''''                #If Trace = 1 Then
''''                    TraceDebug "No 'Shell DocObject View'!", hWndFrame, "E"
''''                #End If
''''            End If
''''        End If ' If GbIsIe7
''''    End If
''''End Function

'--------------------------------------------------------------------------------
' The function is OK, just not necessary

''''Public Function FindActiveBrowserByFrame( _
''''    hWndFrame As Long _
''''    ) As Long
''''
''''    Dim WndChild As Long
''''    Dim IE As WebBrowser
''''
''''    If hWndFrame = 0 Then
''''        #If Trace = 1 Then
''''            TraceDebug "FindActiveBrowserFromFrame: Frame is 0", hWndFrame, "E"
''''        #End If
''''        FindActiveBrowserByFrame = 0
''''    Else
''''        If GbIsIe7 Then
''''            WndChild = FindWindowEx(hWndFrame, 0, "TabWindowClass", vbNullString)
''''            If WndChild <> 0 Then
''''                WndChild = FindWindowEx(WndChild, 0, "Shell DocObject View", vbNullString)
''''                If WndChild <> 0 Then
''''                    WndChild = FindWindowEx(WndChild, 0, "Internet Explorer_Server", vbNullString)
''''                    If WndChild <> 0 Then
''''                        #If Trace = 1 Then
''''                            TraceDebug "Found 'Internet Explorer_Server'!", hWndFrame, "W"
''''                        #End If
''''                        Call GetIEFromHWND(WndChild, IE) 'Get Iwebbrowser2 from Handle
''''                        If Not (IE Is Nothing) Then
''''                            FindActiveBrowserByFrame = WndChild
''''                            #If Trace = 1 Then
''''                                TraceDebug "Got Iwebbrowser2 from Handle, WndChild:" + Str(WndChild), hWndFrame, "W"
''''                            #End If
''''                        Else
''''                            #If Trace = 1 Then
''''                                TraceDebug "IE is not set!", hWndFrame, "E"
''''                            #End If
''''                        End If
''''                    Else
''''                        #If Trace = 1 Then
''''                            TraceDebug "No 'Internet Explorer_Server'!", hWndFrame, "E"
''''                        #End If
''''                    End If
''''                Else
''''                    #If Trace = 1 Then
''''                        TraceDebug "No 'Shell DocObject View'!", hWndFrame, "E"
''''                    #End If
''''                End If
''''            Else
''''                #If Trace = 1 Then
''''                    TraceDebug "No 'TabWindowClass'!", hWndFrame, "E"
''''                #End If
''''            End If
''''        Else ' If GbIsIe7
''''            WndChild = FindWindowEx(hWndFrame, 0, "Shell DocObject View", vbNullString)
''''            If WndChild <> 0 Then
''''                WndChild = FindWindowEx(WndChild, 0, "Internet Explorer_Server", vbNullString)
''''                If WndChild <> 0 Then
''''                    #If Trace = 1 Then
''''                        TraceDebug "Found 'Internet Explorer_Server'!", hWndFrame, "W"
''''                    #End If
''''                    Call GetIEFromHWND(WndChild, IE) 'Get Iwebbrowser2 from Handle
''''                    If Not (IE Is Nothing) Then
''''                        FindActiveBrowserByFrame = WndChild
''''                        #If Trace = 1 Then
''''                            TraceDebug "Got Iwebbrowser2 from Handle, WndChild: " + Str(WndChild), hWndFrame, "W"
''''                        #End If
''''                    Else
''''                        #If Trace = 1 Then
''''                            TraceDebug "IE is not set!", hWndFrame, "E"
''''                        #End If
''''                    End If
''''                Else
''''                    #If Trace = 1 Then
''''                        TraceDebug "No 'Internet Explorer_Server'!", hWndFrame, "E"
''''                    #End If
''''                End If
''''            Else
''''                #If Trace = 1 Then
''''                    TraceDebug "No 'Shell DocObject View'!", hWndFrame, "E"
''''                #End If
''''            End If
''''        End If ' If GbIsIe7
''''    End If
''''End Function


