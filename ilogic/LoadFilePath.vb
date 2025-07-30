Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document
Dim dispName As String = currentDoc.DisplayName
Dim fileName As String = currentDoc.FullFileName
Dim filePath As String = setPathName(fileName, dispName)
Dim docData As PropertySet
Dim docPathName As Object

If currentDoc.PropertySets.PropertySetExists("Path Data") = False Then
    docData = currentDoc.PropertySets.Add("Path Data")
	docPathName = docData.Add(filePath, "Document Path")
	'MsgBox("Access to new created path")	
Else
	docData = currentDoc.PropertySets.Item("Path Data")
	docPathName = docData.Item(1)
	'MsgBox("Access to existing path")
End If

End Sub

Function setPathName (fullFileName As String, DisplayName As String )

   Dim limSup As Integer = Len(fullFileName) 
   Dim limInf As Integer = limSup - Len(DisplayName)
   Dim dirPath As String = fullFileName
   
   dirPath = dirPath.Remove(limInf)
   Return dirPath
   
End Function