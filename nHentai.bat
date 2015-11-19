<!-- : Begin batch script
@ECHO OFF
COLOR DF
TITLE nHentai.Net Downloader
if not exist "./nHentai" (
  mkdir "./nHentai"
)
echo -------------------------------------------------------------------------------
echo *                                                                             *
echo *                          nHentai.Net Downloader                             *
echo *                                                                             *
echo -------------------------------------------------------------------------------
cscript //nologo "%~f0?.wsf" //job:VBS 
echo -------------------------------------------------------------------------------
IF %ERRORLEVEL% EQU 1 EXIT
start "" nHentai.bat
exit
----- Begin wsf script --->
<package>
<job id="VBS">
<script language="VBScript">
Dim nHentaiUrl
nHentaiUrl = InputBox("Please Enter nHentai Gallery Link Example : http://nhentai.net/g/125431/", "nHentai Downloader")

If nHentaiUrl = "" Then
    WScript.Echo ""
	wscript.Quit(1)
Else 
    getImgURL(HTMLDownload(nHentaiUrl))
End If
'--------------------------------
Sub HTTPDownload(strLink)
	nSplit = Split(strLink, "/")
	nFolder = "nHentai/" & nSplit(4)
	strSaveName = Mid(strLink, InStrRev(strLink,"/") + 1, Len(strLink))
	strSaveTo = "./" & nFolder &"/" & strSaveName

	WScript.Echo "* Download URL:" & strLink
	
	Set objHTTP = CreateObject("MSXML2.XMLHTTP")
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
		End With
		imgSize = Round(objStream.Size / 1024)
		WScript.Echo "* File Size : " & imgSize & " KB"
		Info nSplit(4), "* Image Url : " & strLink & " - SaveAs :  " & strSaveTo & " - FileSize : " & imgSize & " KB" 
		objStream.Close
		set objStream = Nothing
	End If

	If objFSO.FileExists(strSaveTo) Then
		WScript.Echo "* Download `" & strSaveTo & "` completed successfuly."
		WScript.Echo "-------------------------------------------------------------------------------"  
	End If
End Sub
'--------------------------------
Dim dict

Function getImgURL(HTMLFile)
    WScript.Echo "-------------------------------------------------------------------------------"
    WScript.Echo "* Gallery image links parsing..."
	WScript.Echo "-------------------------------------------------------------------------------"

    Set Matches = getRegEx(HTMLFile, "data-src=[\""\']([^\""\']+)")
    'Iterate through the Matches collection.
	row = 1
	Set dict = CreateObject("Scripting.Dictionary")
	
    For Each Match in Matches
        'We only want the first match.
		 URL = Match.Value
		 getImgURL = Replace(URL, "data-src=""", "")
		 getImgExt = Mid(URL, InStrRev(URL,".") + 1, Len(URL))
	     getImgURL = "http:" + Replace(getImgURL, "t." & getImgExt, "." & getImgExt)
	     getImgURL = Replace(getImgURL, "://t.", "://i.")	
		 dict.Add row, getImgURL
		 row = row + 1
    Next
    'Clean up
    Set Match = Nothing
    Set RegEx = Nothing
	
	WScript.Echo "* Gallery image links partitioned..."
	WScript.Echo "-------------------------------------------------------------------------------"
	StartDownload dict
	
End Function
'--------------------------------
Function HTMLDownload(strLink)
	With CreateObject("MSXML2.XMLHTTP")
		.open "GET", strLink, False
		.send
		HTMLDownload = .responseText
	End With
End Function
'--------------------------------
Function CreateFolder(NewDir)
	With CreateObject("Scripting.FileSystemObject")
		If Not .FolderExists(NewDir) Then
			.CreateFolder(NewDir)
		End If
	End With
End Function
'--------------------------------
Function getRegEx(HTMLstring, ByRef MatchPattern)
    Set RegEx = New RegExp
    With RegEx
        .Pattern = MatchPattern
        .IgnoreCase = True
        .Global = True
    End With
    Set getRegEx = RegEx.Execute(HTMLstring)
End Function
'--------------------------------
Sub CreateCBZ(sFolder)
    Dim zipFile, SaveName 
	zipFile = sFolder & ".zip"
	
    With CreateObject("Scripting.FileSystemObject")
        zipFile = .GetAbsolutePathName(zipFile)
        sFolder = .GetAbsolutePathName(sFolder)

        With .CreateTextFile(zipFile, True)
            .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, chr(0))
        End With
    End With

    With CreateObject("Shell.Application")
        .NameSpace(zipFile).CopyHere .NameSpace(sFolder).Items

        Do Until .NameSpace(zipFile).Items.Count = _
                 .NameSpace(sFolder).Items.Count
            WScript.Sleep 1000 
        Loop
    End With
    
	With CreateObject("Scripting.FileSystemObject")
        SaveName = Replace(zipFile, ".zip", ".cbz")
		.MoveFile zipFile, SaveName
		.DeleteFolder(sFolder)
	End With
	
End Sub
'--------------------------------
Sub Info(InfoFile, ByVal InfoMsg)
    Const ForAppending = 8
	
	Infos = "./nHentai/"& InfoFile & "/Info.txt"
	
	With CreateObject("Scripting.FileSystemObject")		
		If Not .FileExists(Infos) Then
			Set InfoTxt = .CreateTextFile(Infos)
        End If
	End With
	
    With CreateObject("Scripting.FileSystemObject")
        Set InfoText = .OpenTextFile(Infos, ForAppending)
		InfoText.WriteLine(InfoMsg) 
		InfoText.Close		
	End With
End Sub
'--------------------------------
Function CreateInfo(InfoFile)
	Infos = "./nHentai/"& InfoFile & "/Info.txt"
	
	With CreateObject("Scripting.FileSystemObject")		
		If Not .FileExists(Infos) Then
			Set InfoTxt = .CreateTextFile(Infos)
        End If
	End With
End Function
'--------------------------------
Sub StartDownload(dict)
    WScript.Echo "-------------------------------------------------------------------------------"
	WScript.Echo "*                            DOWNLOAD STARTING!                               *"
	WScript.Echo "-------------------------------------------------------------------------------"
	nSplit = Split(dict(1), "/")
	nFolder = ".\nHentai\" & nSplit(4)
	
	CreateFolder("./" & nFolder)
	WScript.Sleep 1000
	CreateInfo nSplit(4)	
	WScript.Sleep 1000

	Info nSplit(4), "* Gallery URL : " & nHentaiUrl
	
	For Each line in dict
		HTTPDownload dict(line) 
	Next

	WScript.Echo "-------------------------------------------------------------------------------"
	WScript.Echo "*                            DOWNLOAD COMPLETED!                              *"
	WScript.Echo "-------------------------------------------------------------------------------"

	CreateCBZ(nFolder)
	
	WScript.Sleep 3000
	
	intAnswer = Msgbox("Do you want to open?", vbYesNo, "DOWNLOAD COMPLETED!")

	If intAnswer = vbYes Then
		Set WshShell = WScript.CreateObject("WScript.Shell")  
		WshShell.Run(nFolder & ".cbz") 
	Else
		WScript.Echo "Restarting."
	End If
	WScript.Quit
End Sub
</script>
</job>
</package>