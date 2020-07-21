Option Explicit On ' TODO Enable option strict and see what it breaks'
Imports System
Imports System.IO

Module FileWriter

	Sub Main()

		Dim toSearch as String = InputBox("Enter the directory to search", "Directory Path")
		Dim lineNum as Integer = InputBox("Enter the line number", "Line Number")
		Dim toWritePath as String = InputBox("Enter the path for the text to add", "File Path")
		Dim toWrite as String = GetToWrite(toWritePath) ' The user inputs text to add to files via another text file'
		Dim files as String() ' Unintialized array of all the file in toSearch'

		files = GetFiles(toSearch)
		WriteFiles(files, lineNum, toWrite)

	end Sub

	Function GetToWrite(ByVal ltFilePath as String)

		Dim ltFileContents as String = ""
		Dim lReader as StreamReader = New StreamReader(ltFilePath)

		While Not lReader.EndOfStream
			ltFileContents &= lReader.ReadLine ' Copies the entire file in one large string variable'
			ltFileContents &= VbCrLf
		End While

		lReader.close()
		return ltFileContents

	end Function

	Function GetFiles(ByVal toSearch as String)

		Dim directories as String() = Directory.GetDirectories(toSearch) ' Array of all the different directories in specific folder'
		Dim allFiles as String() = {"0"}' Array of all the different files in a given directory.
		Dim files as String() = {"0"} ' Temporary array that has the files within one specific directory'
		Dim numDirectories as Integer = directories.Length
		Dim searches as Integer = 0 ' Loop control variable.'

		for searches = 0 To numDirectories - 1 ' Searches all the directories in the directories array. - 1 b/c otherwise it overflows
			files = Directory.GetFiles(directories(searches), "*.html") ' Looks for .html files'
			if searches = 0 Then
				allFiles = files ' XXX Workaround to having to intialize allFiles. Otherwise VB gives a warning
			else
				allFiles = allFiles.Union(files).ToArray() ' Concatonates the temporary files array to the allFiles array'
			end if
		next

		return allFiles

	end Function

	Sub WriteFiles(ByRef files() As String, ByVal lineNum as String, ByVal toWrite as String)

		Dim count as Integer = 0

		for count = 0 to files.Length - 1 ' Writes a string of text to the file at a certain line. Loop writes to all files'
			AppendAtPosition(files(count), lineNum, toWrite)
		next

	end Sub

	' This method was taken from https://stackoverflow.com/questions/15900225
	Private Sub AppendAtPosition(ByRef ltFilePath As String, ByVal liAppendLine As Integer, ByVal ltToAppend As String)

		Dim ltFileContents As String = ""
		Dim lReader As StreamReader = New StreamReader(ltFilePath)
		Dim liRow As Integer = 0

		While Not lReader.EndOfStream
			ltFileContents &= lReader.ReadLine
			If liRow = liAppendLine Then ' Appends text at the specified line'
				ltFileContents &= ltToAppend
			end if

			ltFileContents &= VbCrLf
  			liRow += 1
		End While

		lReader.close()

		File.WriteAllText(ltFilePath, ltFileContents)

	End Sub

end Module
