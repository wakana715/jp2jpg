Const strFileList = "fileList.txt"
Const strFromExt  = "jp2"
Const strToExt    = "jpg"
Const strCmd      = "cmd /c"
Const strDir      = "dir /b /a-d /s"
Const strConvert  = "C:\ImageMagick\convert.exe -resize x2048"
Dim objShell, objFileSystem, objFileList
Dim strFrom, strTo, strRun, intPos, strTmp1, strTmp2, strTmp3
Set objShell      = CreateObject("WScript.Shell")
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
strRun = strCmd & " " & strDir & " *." & strFromExt & ">" & strFileList
WScript.Echo strRun
objShell.Run strRun, 0, True
Set objFileList = ObjFileSystem.OpenTextFile(strFileList, 1, False)
Do Until objFileList.AtEndOfStream
	strFrom = objFileList.ReadLine
	strTo   = Left(strFrom, Len(strFrom) - Len(strFromExt)) & strToExt
	intPos  = InStrRev(strFrom, "\", -1, 1)
	If intPos > 0 Then
		strTmp1 = Left(strFrom, intPos) & strToExt
		On Error Resume Next
		objFileSystem.CreateFolder(strTmp1)
		On Error Goto 0
		strTmp2 = Mid(strFrom, intPos + 1)
		strTmp3 = Left(strTmp2, Len(strTmp2) - Len(strFromExt))
		strTo = strTmp1 & "\" & strTmp3 & strToExt
	End If
	strRun  = strCmd & " " & strConvert & " " & strFrom & " " & strTo
	WScript.Echo strRun
	objShell.Run strRun, 0, True
Loop
objFileList.Close
