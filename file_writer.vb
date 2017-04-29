Option Explicit On ' TODO Enable option strict and see what it breaks'
Imports System
Imports System.IO

Module FileWriter

	Sub Main() ' TODO make private'

		Dim toSearch as String = "D:\Projects\Github\file_writer\test_folders" ' TODO consider defaulting to current directory if no input'
		Dim lineNum as Integer = 5 ' TODO allow for user input on the command line'
		Dim toWrite as String = "Test" & VbCrLf & "Test 2" ' TODO same as above'
		Dim files as String()

		files = GetFiles(toSearch)
		WriteFiles(files, lineNum, toWrite)

	end Sub

	Function GetFiles(toSearch)

		Dim directories as String() = Directory.GetDirectories(Convert.toString(toSearch)) ' Array of all the different directories in specific folder' TODO find workaround to Convert.toString
		Dim allFiles as String() ' Array of all the different files in a given directory. TODO remove need to initialize
		Dim files as String() = {"0"} ' Temporary array that has the files within one specific directory'
		Dim numDirectories as Integer = directories.Length
		Dim searches as Integer = 0 ' Loop control variable.'


		for searches = 0 To numDirectories - 1 ' Searches all the directories in the directories array. - 1 b/c otherwise it overflows
			files = Directory.GetFiles(directories(searches), "*.html") ' Looks for .html files'
			if searches = 0 Then
				allFiles = files ' FIXME don't let the compiler think that it isn't actually initialzed here.
			else
				allFiles = allFiles.Union(files).ToArray() ' Concatonates the temporary files array to the allFiles array'``
			end if
		next

		return allFiles

	end Function

	Sub WriteFiles(files, lineNum, toWrite)

		Dim count as Integer = 0

		for count = 0 to files.Length - 1 ' Writes a string of text to the file at a certain line. Loop writes to all files'
			AppendAtPosition(files(count), lineNum, toWrite)
		next

	end Sub

	Private Sub AppendAtPosition(ByVal ltFilePath As String, ByVal liAppendLine As Integer, ByVal ltAppendLine As String) ' Prewritten code. Does the file writting'
		'TODO when it finds the right line, save the rest of the file so it isn't deleted when text is appended. Then reappend the saved text
		Dim ltFileContents As String = ""
		Dim lReader As StreamReader = New StreamReader(ltFilePath)
		Dim liRow As Integer = 0

		While Not lReader.EndOfStream ' Searches for line until end of file'
			ltFileContents &= lReader.ReadLine
			If liRow = liAppendLine Then ' Appends text at the specified line'
				ltFileContents &= ltAppendLine
				Exit While ' FIXME remove "spaghetti" code'
			End If
  			liRow += 1
		End While

		If liAppendLine >= liRow Then
			ltFileContents &= ltAppendLine
		End If


		File.WriteAllText(ltFilePath, ltFileContents) ' FIXME this lime here crashes it. Supposedly it is being accessed by another process. (Something in the code?)'

	End Sub

end Module
