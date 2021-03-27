Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("file_bfo", 1)
Dim fso2, file2
Set fso2 = CreateObject("Scripting.FileSystemObject")
Set file2 = fso2.CreateTextFile("file_bfo.out", True)
Do Until file.AtEndOfStream
	line = file.Readline
	if Len(line) > 0 Then
		line = LTrim(line)
		if (InStr(line,"ANS8000I") = 0) Then
			file2.WriteLine(line)
		end if
	end if
Loop
file2.Close
