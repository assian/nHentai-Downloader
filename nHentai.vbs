if WScript.Arguments.Count < 1 then
  MsgBox "Usage: iGet.vbs <list file>"
  WScript.Quit
end if
'/////////////////////////////////////////////////////////////////////////
Sub HTTPDownload(strLink)
nFolder = Split(strLink, "/")
strSaveName = Mid(strLink, InStrRev(strLink,"/") + 1, Len(strLink))
strSaveTo = "./" & nFolder(4) &"/" & strSaveName
CreateFolder("./" & nFolder(4))
WScript.Echo "-------------"
WScript.Echo "Download: " & strLink
WScript.Echo "Save to:  " & strSaveTo

' Create an HTTP object
Set objHTTP = CreateObject("MSXML2.XMLHTTP")

' Download the specified URL
'xmlhttp.Open "GET", strURL, false, "User", "Password"
objHTTP.open "GET", strLink, False
objHTTP.send
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strSaveTo) Then
  objFSO.DeleteFile(strSaveTo)
End If

If objHTTP.Status = 200 Then
  Dim objStream
  Set objStream = CreateObject("ADODB.Stream")
  With objStream
    .Type = 1 'adTypeBinary
    .Open
    .Write objHTTP.responseBody
    .SaveToFile strSaveTo
    .Close
  End With
  set objStream = Nothing
End If

If objFSO.FileExists(strSaveTo) Then
  WScript.Echo "Download `" & strSaveName & "` completed successfuly."
End If
End Sub
'/////////////////////////////////////////////////////////////////////////
Function getImgTagURL(HTMLFile)
    WScript.Echo "Gallery image links parsing..."
    Set fso = CreateObject("Scripting.FileSystemObject")
    HTMLstring = fso.OpenTextFile(HTMLFile).ReadAll
    Set RegEx = New RegExp
    With RegEx
        .Pattern = "data-src=[\""\']([^\""\']+)"
        .IgnoreCase = True
        .Global = True
    End With

    Set Matches = RegEx.Execute(HTMLstring)
    'Iterate through the Matches collection.
    URL = ""
	Set objFile = fso.OpenTextFile(HTMLFile, 2)
    For Each Match in Matches
        'We only want the first match.
		 URL = Match.Value
		 getImgTagURL = Replace(URL, "data-src=""", "")
	     getImgTagURL = "http:" + Replace(getImgTagURL, "t.jpg", ".jpg")
	     getImgTagURL = Replace(getImgTagURL, "://t.", "://i.")		 
		 objFile.WriteLine getImgTagURL
    Next
	objFile.Close
    'Clean up
    Set Match = Nothing
    Set RegEx = Nothing
	WScript.Echo "Gallery image links partitioned..."
	StartDownload(HTMLFile)
    ' src=" is hanging on the front, so we will replace it with nothing

End Function
'/////////////////////////////////////////////////////////////////////////
Sub HTMLDownload(strLink)
WScript.Echo "Gallery information downloading..."
Set objHTTP = CreateObject("MSXML2.XMLHTTP")
objHTTP.open "GET", strLink, False
objHTTP.send
strSaveTo = "./temp.txt"
If objHTTP.Status = 200 Then
  Dim objStream
  Set objStream = CreateObject("ADODB.Stream")
  With objStream
    .Type = 1 'adTypeBinary
    .Open
    .Write objHTTP.responseBody
    .SaveToFile strSaveTo
    .Close
  End With
End If
WScript.Echo "Gallery information downloaded."
getImgTagURL("./temp.txt")
End Sub
'/////////////////////////////////////////////////////////////////////////
Sub CreateFolder(NewDir)
dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")
If Not oFSO.FolderExists(NewDir) Then
Set objFolder = oFSO.CreateFolder(NewDir)
End If
set oFSO=nothing
End Sub
'/////////////////////////////////////////////////////////////////////////
Sub StartDownload(DownloadList)
WScript.Echo "Images download starting..."
Set fso = CreateObject("Scripting.FileSystemObject")
Set dict = CreateObject("Scripting.Dictionary")
Set file = fso.OpenTextFile (DownloadList, 1)
row = 0
Do Until file.AtEndOfStream
  line = file.Readline
  dict.Add row, line
  row = row + 1
Loop

file.Close
WScript.Echo "Images download started."
'Loop over it
For Each line in dict.Items
  HTTPDownload line
Next
' Done
Set obj = CreateObject("Scripting.FileSystemObject")
obj.DeleteFile("./temp.txt")
WScript.Quit
End Sub
HTMLDownload(WScript.Arguments(0))