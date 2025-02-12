Const strFileList = "fileList.txt"
Const strFromInfo = "fromInfo.txt"
Const strFromExt  = "jp2"
Const strToExt    = "jpg"
Const strCmd      = "cmd /c"
Const strDir      = "dir /b /a-d /s"
Const strIdentify = "C:\ImageMagick\identify"
Const strConvert1 = "C:\ImageMagick\convert.exe"
Const strConvert2 = "C:\ImageMagick\convert.exe -resize x2048"
Const lngResizeHeight =                                  2048
Dim objShell, objFileSystem, objFileList, objIdentify
Dim strFrom, strTo, strRun, strConvert, strTmp1, strTmp2, strTmp3
Dim intPos, lngIdentifyHeight
Set objShell      = CreateObject("WScript.Shell")
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
strRun = strCmd & " " & strDir & " *." & strFromExt & ">" & strFileList
WScript.Echo strRun
objShell.Run strRun, 0, True
Set objFileList = ObjFileSystem.OpenTextFile(strFileList, 1, 0) '1:ForReading, 0:ASC
Do Until objFileList.AtEndOfStream
	strFrom = objFileList.ReadLine
	strRun  = strCmd & " " & strIdentify & " " & strFrom & ">" & strFromInfo
	WScript.Echo strRun
	objShell.Run strRun, 0, True
	Set objIdentify = ObjFileSystem.OpenTextFile(strFromInfo, 1, 0) '1:ForReading, 0:ASC
	strTmp1 = objIdentify.ReadAll
	objIdentify.Close
	strConvert = strConvert2
	strTmp1 = Mid(strTmp1, Len(strFromInfo) + 2)
	intPos = InStr(strTmp1, " ")
	If intPos > 0 Then
		strTmp1 = Mid(strTmp1, intPos + 1)
		intPos = InStr(strTmp1, "x")
		If intPos > 0 Then
			strTmp1  = Mid(strTmp1, intPos + 1)
			intPos = InStr(strTmp1, " ")
			If intPos > 0 Then
				strTmp2  = Mid(strTmp1, 1, intPos - 1)
				If Len(strTmp2) > 0 Then
					lngIdentifyHeight = CLng(strTmp2)
					If lngIdentifyHeight < lngResizeHeight Then
						strConvert = strConvert1
					End If
				End If
			End If
		End If
	End If
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
