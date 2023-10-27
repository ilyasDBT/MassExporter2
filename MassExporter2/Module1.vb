Imports Inventor
Module Module1

	Sub Main(ByVal args() As String)
		Dim versionNumber As String = "0.0.5"
		Console.WriteLine("MassExporter2 version:" & versionNumber)
		Dim invApp As Inventor.Application = CreateObject("Inventor.Application")
		invApp.SilentOperation = True
		'invApp.Visible = True
		'Console.WriteLine("Hello World!")
		Dim pathsFilePath As String = args(0)
		Dim outputFilePath As String = args(1)

		'Dim pathsFilePath As String = "C:\Users\Ilyas\Desktop\myfiles3.csv"
		'Dim outputFilePath As String = "C:\Users\Ilyas\Desktop\mass.csv"

		Dim outputFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(outputFilePath, False)
		outputFile.WriteLine("MassExporter2 version:" & versionNumber)

		Dim reader As IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(pathsFilePath)
		Dim pathsList As New List(Of String)
		While (reader.Peek() <> -1)
			pathsList.Add(reader.ReadLine())
		End While

		Dim totalFiles As Integer = pathsList.Count
		'Logger.Trace("Number of Files: " & totalFiles)

		Dim oDocument As Document
		Dim mass As String
		Dim partNumber As String
		Dim count As Integer = 0
		For Each filePath As String In pathsList
			count += 1
			Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
			If Not System.IO.File.Exists(filePath) Then
				outputFile.WriteLine(count & "," & fileName & ",Error file not found")
				Continue For
			End If

			Try
				oDocument = invApp.Documents.Open(filePath, False)
			Catch
				outputFile.WriteLine(count & "," & fileName & ",Error opening file")
				'Logger.Trace(count & "/" & totalFiles & " : " & fileName & ", Error")
				Continue For
			End Try

			Try
				mass = CDbl(Math.Round(oDocument.ComponentDefinition.MassProperties.Mass, 3)).ToString("0.000") ' & "kg"
				partNumber = oDocument.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
				outputFile.WriteLine(count & "," & partNumber & "," & mass)
				'Logger.Trace(count & "/" & totalFiles & " : " & fileName & ", " & mass)
			Catch
				outputFile.WriteLine(count & "," & fileName & ",Error getting properties")
				'Logger.Trace(count & "/" & totalFiles & " : " & fileName & ", Error")
			Finally
				oDocument.Close(True)
			End Try
		Next

		outputFile.Close()
		invApp.Quit()
	End Sub

End Module
