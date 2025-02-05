Const strFileList = "fileList.txt"
Const strFromExt  = "jpg"
Const strToExt    = "png"
Const strCmd      = "cmd /c"
Const strDir      = "dir /b /a-d /s"
Const strConvert  = "C:\ImageMagick\convert.exe -resize x2048"
Dim objShell, objFileSystem, objFileList
Dim strFrom, strTo, strRun
Set objShell      = CreateObject("WScript.Shell")
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
strRun = strCmd & " " & strDir & " *." & strFromExt & ">" & strFileList
WScript.Echo strRun
objShell.Run strRun, 0, True
Set objFileList = ObjFileSystem.OpenTextFile(strFileList, 1, False)
Do Until objFileList.AtEndOfStream
	strFrom = objFileList.ReadLine
	strTo   = Left(strFrom, Len(strFrom) - Len(strFromExt))
	strRun  = strCmd & " " & strConvert & " " & strFrom & " " & strTo & strToExt
	WScript.Echo strRun
	objShell.Run strRun, 0, True
Loop
objFileList.Close
