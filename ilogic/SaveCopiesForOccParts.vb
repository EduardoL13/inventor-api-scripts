Sub Main ()
' This routines make a "Save As" to all part type occurrences within the assembly and saves them into the path given by the user with a sufix given by the user.
	
Dim currentDoc As AssemblyDocument = ThisDoc.Document
Dim fileName As String

nameCompRoute = InputBox("Enter the path for the location in which occurrences will be saved", "Files for Ocurrences", "")

sufixName = InputBox("Enter the desired sufix", "Files for Ocurrences", "Default Entry")

Dim sufix As String = "_" & sufixName & "_"

For Each compOcc As ComponentOccurrence In currentDoc.ComponentDefinition.Occurrences
	
	Dim oNewDoc As PartDocument = compOcc.Definition.Document
	fileName = setPathName(compOcc.Name)
	ThisApplication.SilentOperation = True
	oNewDoc.SaveAs(nameCompRoute & "\" & fileName & sufix & ".ipt", False) 
	ThisApplication.SilentOperation = False
    
Next
End Sub

Function setPathName (fullFileName As String )

   Dim limSup As Integer = Len(fullFileName) 
   Dim limInf As Integer = limSup - 2
   Dim dirPath As String = fullFileName
   
   dirPath = dirPath.Remove(limInf)
   Return dirPath
   
End Function
