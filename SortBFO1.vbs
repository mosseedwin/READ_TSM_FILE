Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("file_bfo.out", 1)
Dim fso1, file1
Set fso1 = CreateObject("Scripting.FileSystemObject")
Set file1 = fso1.CreateTextFile("file_bfo.1.out", True)
Dim fso2, file2
Set fso2 = CreateObject("Scripting.FileSystemObject")
Set file2 = fso2.CreateTextFile("file_bfo.2.out", True)
Dim fso3, file3
Set fso3 = CreateObject("Scripting.FileSystemObject")
Set file3 = fso3.CreateTextFile("file_bfo.3.out", True)
Dim fso4, file4
Set fso4 = CreateObject("Scripting.FileSystemObject")
Set file4 = fso4.CreateTextFile("file_bfo.4.out", True)
Dim fso5, file5
Set fso5 = CreateObject("Scripting.FileSystemObject")
Set file5 = fso5.CreateTextFile("file_bfo.5.out", True)
Do Until file.AtEndOfStream
	objectId = file.Readline
	if Len(objectId) > 0 Then
		if (InStr(objectId,"Bitfile Object: ") > 0) Then
			line = file.Readline
			if (InStr(line,"Bitfile Object NOT found") > 0) Then
				file1.WriteLine(objectId)
			elseif (InStr(line,"**Archival") > 0) Then
				file.SkipLine
				file2.Write(objectId)
				file2.Write(vbTab)
				file2.Write(file.Readline)
				file2.Write(vbTab)
				file2.Write(file.Readline)
				file2.Write(vbNewLine)
			elseif (InStr(line,"**Super-bitfile") > 0) Then
				Do Until False
					line = file.Readline
					if (InStr(line,"**Archival") = 1) Then
						Exit Do
					end if
				Loop
				file.SkipLine
				file3.Write(objectId)
				file3.Write(vbTab)
				file3.Write(file.Readline)
				file3.Write(vbTab)
				file3.Write(file.Readline)
				file3.Write(vbNewLine)
			elseif (InStr(line,"**Sub-bitfile") > 0) Then
				file4.Write(objectId)
				file4.Write(vbTab)
				file4.Write(file.Readline)
				file4.Write(vbNewLine)
			end if
		end if
	end if
Loop
file2.Close
