Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long



Public Sub Launch_VBAExtractor()

frmVBAGitExtractor.Show

End Sub



Sub DownloadFile()

Dim myURL As String
myURL = "https://github.com/badkatro/AssignDocuments/blob/master/scripts/Main.vba"

'Dim WinHttpReq As Object
Dim WinHttpReq As XMLHTTP
'Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
Set WinHttpReq = New XMLHTTP

WinHttpReq.Open "GET", myURL, False, "badkatro", "Kocutel924"
WinHttpReq.send

Do While WinHttpReq.readyState <> 4
    DoEvents
Loop

myURL = WinHttpReq.responseBody

If WinHttpReq.Status = 200 Then
    
    Dim fso As New Scripting.FileSystemObject
    
    Dim ostream As TextStream
    
    Set ostream = fso.CreateTextFile("D:\Main.bas", True)
    'Set oStream = CreateObject("ADODB.Stream")
    
    'oStream.Open
    ostream.write (WinHttpReq.responseText)
    'ostream.Type = 1
    'ostream.write WinHttpReq.responseText
    'ostream.SaveToFile "D:\Main.bas", 2 ' 1 = no overwrite, 2 = overwrite
    ostream.Close
End If

Set WinHttpReq = Nothing

End Sub

Sub Test_GitHub_Iexplorer_Dwn()


Call GitHub_Download_Repo("C:\Users\" & Environ("Username") & "\Documents\VBA\TestDown", "https://github.com/badkatro/AssignDocuments")


End Sub


Function GitHub_Download_Repo(StageFolder As String, RepoAddress As String) As Boolean


Dim iexplorer As New InternetExplorer


Debug.Print "IExplorer: Navigating to: " & RepoAddress & "..."

iexplorer.Visible = False

iexplorer.Navigate (RepoAddress)


Do While iexplorer.Busy
    DoEvents
Loop

If iexplorer.readyState = READYSTATE_COMPLETE Then
    Debug.Print "InterExplorer: ReadyState = Complete: " & CStr(iexplorer.readyState = READYSTATE_COMPLETE)
Else
    Debug.Print "InterExplorer: ReadyState = Complete: " & CStr(iexplorer.readyState = READYSTATE_COMPLETE)
End If


' HTML doc try... prob not
Dim topHtml As HTMLDocument
Set topHtml = iexplorer.Document
'iexplorer.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT

' enumerate all props of link41, which links to a file, so as to find one usable... (in finding all usable links)
' Found: (prop num, prop localname, prop value)
' prop 35: tabindex - 1, prop 121: class - accessibility-aid js-skip-to-content, prop 160: href - #start-of-content
'For i = 1 To topHtml.Links.Item(, 41).Attributes.Length - 1
'    If topHtml.Links.Item(, 41).Attributes.Item(i).Value <> "" Then
'        Debug.Print i & ": "; topHtml.Links.Item(, 41).Attributes.Item(i).localName & " - " & topHtml.Links.Item(, 41).Attributes.Item(i).Value
'    End If
'Next i

' enumerate all links in doc to get those with classname = "accessibility-aid js-skip-to-content"
' to see if these aret the links we wish

Dim lnks As IHTMLElementCollection

Set lnks = topHtml.Links

Dim lnk As HTMLAnchorElement

Dim dwn As Long     ' result of downloading file

Dim lnkFdrs As New Collection    ' links to folders, need to navigate to 'em


Dim LocalStageF As String
LocalStageF = StageFolder & "\" & Split(RepoAddress, "/")(UBound(Split(RepoAddress, "/")))

' and create base folder for repo if not already
If Dir(LocalStageF, vbDirectory) = "" Then MkDir (LocalStageF)


For Each lnk In lnks
    
    If lnk.ClassName = "js-navigation-open" Then
        
        Debug.Print "Link " & lnk.textContent & ", with href = " & lnk.href & ", mimeType = " & lnk.mimeType
        
        If InStr(1, lnk.mimeType, "File") > 0 And InStr(1, lnk.mimeType, "/") = 0 Then
            
            Dim NFilename As String
            
            NFilename = Split(lnk.href, "/")(UBound(Split(lnk.href, "/")))
            
            Dim LocalFilePath As String
            
            LocalFilePath = LocalStageF & NFilename
            
            Debug.Print "Will download " & NFilename & " to " & LocalFilePath
            
            dwn = URLDownloadToFile(0, lnk.href, LocalFileName, 0, 0)
            
            'PauseForSeconds (0.6)
            
            Debug.Print "Download of " & Split(lnk.href, "/")(UBound(Split(lnk.href, "/"))) & " : " & IIf(dwn = 0, "Success", "Failure") & vbCr
        
        Else
            
            Debug.Print "Found link to folder " & Split(lnk.href, "/")(UBound(Split(lnk.href, "/"))) & _
                ", Saving for later..."
            
            lnkFdrs.add lnk.href, lnk.textContent
            
            MkDir (LocalStageF & "\" & Split(lnk.href, "/")(UBound(Split(lnk.href, "/"))))
            
            'PauseForSeconds 0.7
            
        End If
        
    End If
    
Next lnk



'
If lnkFdrs.Count > 0 Then
    
    Debug.Print vbCr & "Will descend into " & lnkFdrs.Count & " directories"
    
    For i = 1 To lnkFdrs.Count
                
        Debug.Print "Navigating to " & lnkFdrs.Item(i)
        iexplorer.Navigate (lnkFdrs.Item(i))
        
        Do While iexplorer.readyState <> READYSTATE_COMPLETE
            DoEvents
        Loop

        If iexplorer.readyState = READYSTATE_COMPLETE Then
            Debug.Print "InterExplorer: ReadyState = Complete: " & CStr(iexplorer.readyState = READYSTATE_COMPLETE)
        End If
        
        
        ' AND REPEAT
        
        Set topHtml = iexplorer.Document
        
        Set lnks = topHtml.Links
        
        
        For Each lnk In lnks
            
            If lnk.ClassName = "js-directory-link js-navigation-open" Then
                Debug.Print "Link " & lnk.textContent & ", with href = " & lnk.href & ", mimeType = " & lnk.mimeType
                
                If InStr(1, lnk.mimeType, "File") > 0 And InStr(1, lnk.mimeType, "/") = 0 Then
                    LocalFileName = "C:\Users\" & Environ("Username") & "\Documents\VBA\TestDown\AssignDocuments\" & Split(lnk.href, "/")(UBound(Split(lnk.href, "/")))
                    Debug.Print "Will download " & Split(lnk.href, "/")(UBound(Split(lnk.href, "/"))) & " to " & LocalFileName
                    dwn = URLDownloadToFile(0, lnk.href, LocalFileName, 0, 0)
                    Debug.Print "Download of " & Split(lnk.href, "/")(UBound(Split(lnk.href, "/"))) & " : " & IIf(dwn = 0, "Success", "Failure") & vbCr
                    
                    PauseForSeconds (0.6)
                    
                End If
                
            End If
        Next lnk
        
        
    Next i
    
End If

'Debug.Print "IExplorer: Navigating to: " & iexplorer.LocationURL & "..."
'
'iexplorer.Navigate ("https://github.com/badkatro/AssignDocuments/tree/master/scripts")
'
'Do While iexplorer.Busy
'    DoEvents
'Loop
'
'If iexplorer.ReadyState = READYSTATE_COMPLETE Then
'    Debug.Print "InterExplorer: ReadyState = Complete: " & CStr(iexplorer.ReadyState = READYSTATE_COMPLETE) & " " & iexplorer.LocationURL
'End If
'
''Dim myReq As New WinHttp.WinHttpRequest
'
'Set topHtml = iexplorer.Document
'
'If topHtml.anchors.Length > 0 Then
'    Dim addr As String
'
'    For i = 1 To topHtml.anchors.Length
'        addr = topHtml.anchors.Item(i).href
'    Next i
'End If

Debug.Print "All done, Quitting."

Set iexplorer = Nothing

End Function





