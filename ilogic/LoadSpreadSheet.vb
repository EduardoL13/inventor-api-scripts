Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document
Dim DSData As PropertySet

If currentDoc.PropertySets.PropertySetExists("Spreadsheet Document") = False Then
    DSData = currentDoc.PropertySets.Add("Spreadsheet Document")
	docNameString = InputBox("Enter the Excel document name", "ExcelFile1", "Excel Document Name")
	fullDocNameString = docNameString & ".xlsx"
	docName = DSData.Add(fullDocNameString, "File Name")
	MsgBox("File Name:" & docName & " has been added")
Else
	DSData = currentDoc.PropertySets.Item("Spreadsheet Document")
	docName = DSData.Item(1)
	'MsgBox("Existing worksheet will be used")
End If
MsgBox(docName.Value)

End Sub