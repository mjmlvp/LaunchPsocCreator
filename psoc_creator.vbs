dim GoodFiletypes
GoodFiletypes = Array("cywrk","cyprj")
PsocCreatorVersionSearchstring = "<WriteAppVersionLastSavedWith"
PsocCreatorVersionAttributeSearchstring = "v="""

call Main

Sub Main()
	Dim filesys
	Dim strVersion
	gotVersion = False

	Set filesys = CreateObject("Scripting.FileSystemObject")

	' Store the arguments in a variable:
	Set objArgs = Wscript.Arguments
	
	Set objTextStream = Nothing
	
	If (objArgs.Count > 0) Then
	If (filesys.FileExists(objArgs.Item(0))) Then
	
		'WScript.Echo "opening " & objArgs.Item(0)
		
		Set objFile = filesys.GetFile(objArgs.Item(0))
		
		extension = filesys.GetExtensionName(objFile.path)
		
		rem WScript.Echo "Extension: " & extension
		
		For Each filetype in GoodFiletypes
			If extension = filetype Then
				rem WScript.Echo "Good: Extension = " & filetype
				
				Set objTextStream = objFile.OpenAsTextStream
				
				Exit For
				
			end If
		Next
		
		If Not objTextStream Is Nothing Then

			'WScript.Echo "Reading from file '" & objFile.Name & "'"
			
			Dim fileText
			fileText = objTextStream.ReadAll()
			
			'WScript.Echo Len(fileText)
			
			pos = InStr(fileText,PsocCreatorVersionSearchstring)
			'If substr found
			If pos > 0 Then
				fileText = Right(fileText, 1 + Len(fileText) - pos - Len(PsocCreatorVersionSearchstring))
				
				fileText = Left(fileText,InStr(fileText,"/>") - Len("/>"))
				'WScript.Echo "$" & fileText & "$"
				
				pos = Instr(fileText,PsocCreatorVersionAttributeSearchstring)
				If pos > 0 Then
					fileText = Right(fileText, 1 + Len(fileText) - pos - Len(PsocCreatorVersionAttributeSearchstring))
				
					strVersion = Left(fileText,InStr(fileText,"""") - Len(""""))
					
					gotVersion = True

					'WScript.Echo "$" & fileText & "$"
				End If
	
				
				
				
			Else
				WScript.Echo "not found!!"
			End If
			
			
			'fileText = Left(fileText,InStr(fileText,"/>"))
			
			

		end If
	end If
	end If
	
	If gotVersion = False Then
		WScript.Echo "Cannot determine what version of PsocCreator to use" & vbCrLf & "Which do you want?"
		
		strFolderSelected = SelectFolder(ScriptPath())

		If strFolderSelected = vbNull Then
			WScript.Echo "Cancelled"
		Else
			'WScript.Echo "Selected Folder: """ & strFolderSelected & """"
			
			strVersion = filesys.GetFileName(strFolderSelected)
			
			gotVersion = True
		End If
		
		
	End If	
	
	If gotVersion = True Then
		'WScript.Echo strVersion
	
		Set objScriptFolder = filesys.GetFolder(ScriptPath())
		Set arrSubFolders = objScriptFolder.SubFolders
		
		' Display all directories
		For Each dir in arrSubFolders
			If InStr(strVersion,dir.Name) = 1 Then
				
				argline = ""
				
				For i = 0 to objArgs.Count - 1
					argline = argline & objArgs.Item(i)
					
					If(i <> objArgs.Count - 1) Then
						argline = argline & " "
					End If
				Next
				
			
				binfolder = dir.Path & "\PSoC Creator\bin"
			
				If objArgs.Count > 0 Then
					runfolder = filesys.GetParentFolderName(objArgs.Item(0))
				Else
					runfolder = binfolder
				End If

				'WScript.Echo argline
				
				Set objShell = CreateObject("Shell.Application")
				
				runline = "objShell.ShellExecute """ & binFolder & "\psoc_creator.exe"", " & vbNewLine & """" & argline & """," & vbNewLine & """" & runfolder & """," & vbNewLine & """open"", 1"
				
				'assume single argument holding 
				If argline <> "" Then
					argline = """" & argline & """"
				End If
				
				WScript.Echo "argline:"&argline
							
				WScript.Echo "shall run: """ & dir.Path & """" & vbNewLine & "using vbscript line: " & runline
				
				objShell.ShellExecute """" & binFolder & "\psoc_creator.exe""", argline, """" & runfolder & """", "open", 1
				
				Exit For
			End If

		Next
	

	End If
	
End Sub


Function ScriptPath()
	Dim filesys
	Set filesys = CreateObject("Scripting.FileSystemObject")
	varPathCurrent = filesys.GetParentFolderName(WScript.ScriptFullName)
	ScriptPath = varPathCurrent
End Function



Function SelectFolder( myStartFolder )

    ' Standard housekeeping
    Dim objFolder, objItem, objShell
    
    ' Custom error handling
    On Error Resume Next
    SelectFolder = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, myStartFolder )

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function
