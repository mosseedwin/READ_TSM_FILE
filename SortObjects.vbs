Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("file1.txt", 1)
Dim fso2, file2
Set fso2 = CreateObject("Scripting.FileSystemObject")
Set file2 = fso2.CreateTextFile("file1.out", True)
Do Until file.AtEndOfStream
	line = file.Readline
	if Len(line)<>0 Then
		index0 = InStr(line,": ")
		index1 = Len(line)-index0
		file2.Write(LTrim(Right(line,index1)))
		file2.Write(vbTab)
	else
		file2.Write(vbNewLine)
	end if
Loop
file2.Close
