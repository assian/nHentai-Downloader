@ECHO OFF
TITLE nHentai Gallery Downloader v0.2 Beta
:wscript.echo InputBox("Enter nHentai Gallery Link")
findstr "^:" "%~sf0">temp.vbs & for /f "delims=" %%N in ('cscript //nologo temp.vbs') do set link=%%N & del temp.vbs
cscript //nologo nHentai.vbs %link%
pause