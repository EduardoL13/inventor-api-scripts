Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document
Dim DSData As PropertySet

If currentDoc.PropertySets.PropertySetExists("DS Data") = False Then
    DSData = currentDoc.PropertySets.Add("DS Data")
	nameTabDS = InputBox("Enter the target worksheet name", "Worksheet1", "Worksheet Name")
	docPathName = DSData.Add(nameTabDS, "Worksheet")
	MsgBox("Worksheet:" & nameTabDS & " has been added")
Else
	DSData = currentDoc.PropertySets.Item("DS Data")
	docPathName = DSData.Item(1)
	'MsgBox("Existing worksheet will be used")
End If
MsgBox(docPathName.Value)

End Sub