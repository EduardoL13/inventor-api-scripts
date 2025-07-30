Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document

Dim sFile As String
Dim tabDS As String
Dim outputForDS = getDS(currentDoc)
sFile = outputForDS(0)
tabDS = outputForDS(1)
	
Dim lastDataValue As Integer = findLastDataRow(sFile, tabDS) '4 'Definición de Range 

'-----
setPropertiesAndMat(assemDef, sFile, tabDS, lastDataValue) 
	
End Sub	


Function findLastDataRow(file As String, tab As String)
	Dim cellVal As Object 
	Dim range As Integer = 1000
	Dim lastDataRow As Integer = 0
    For rowNum As Integer = 1 To range
		cellVal = GoExcel.CellValue(file, tab, "A" & rowNum)
	    If cellVal Is Nothing  Then '(cellVal Is Nothing OrElse String.IsNullOrEmpty(cellVal.ToString()))
			If (rowNum > range) Then Exit For
	    Else
			lastDataRow = rowNum + 1
		End If
	Next
	Return lastDataRow
End Function

Sub setPropertiesAndMat(assembComp As AssemblyDocument, file As String, tab As String, lastValue As Integer) 
For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences
    If compOcc.Suppressed Then
	
	Else
	    For rowCounter=2 To lastValue
		Dim nameOccDS As String = GoExcel.CellValue(file, tab, "A" & rowCounter) 
		'MsgBox(compOcc.Name)
		
		'Make sure to put here below the corresponding column letter to the property you want to assign
		
		colPartNo = "M" 'Part Number
		colStockNumber = "N" 'Stock Number
		colDescription = "D" 'Description
		colMaterial = "E" 'Material 
		
		If compOcc.Name = nameOccDS Then
		
        	iProperties.Expression(compOcc.Name, "Project", "Part Number") = GoExcel.CellValue(file, tab, colPartNo & rowCounter)
			iProperties.Expression(compOcc.Name, "Project", "Stock Number") = GoExcel.CellValue(file, tab, colStockNumber & rowCounter)
	    	iProperties.Expression(compOcc.Name, "Project", "Description") = GoExcel.CellValue(file, tab, colDescription & rowCounter)
	    	iProperties.MaterialOfComponent(compOcc.Name) = GoExcel.CellValue(file, tab, colMaterial & rowCounter) ' value for this property should be given the exact way Inventor assigns it when assigning the material manually
		End If
	Next
	End If
	Next

End Sub

Function LoadDSComms(defDoc As Object, stringToLoad As String)
    If stringToLoad = "worksheet" Then
        iLogicVb.RunRule("LoadWorkSheetName")
		Dim oPropSet As PropertySet = defDoc.PropertySets.Item("DS Data")
		Dim tabDS As String = oPropSet.Item(1).Value
    	Return tabDS
	Else If stringToLoad = "path"
		iLogicVb.RunRule("LoadFilePath")
		Dim oPropSet As PropertySet = defDoc.PropertySets.Item("Path Data")
		Dim pathName As String = oPropSet.Item(1).Value
    	Return pathName
	Else If stringToLoad = "document"
		iLogicVb.RunRule("LoadSpreadSheet")
		Dim oPropSet As PropertySet = defDoc.PropertySets.Item("Spreadsheet Document")
		Dim documentSpreadName As String = oPropSet.Item(1).Value
    	Return documentSpreadName
	Else
		MsgBox("Input inválido")
	End If
End Function

Function getDS(activeDoc As Object)
	
	Dim ruta As String = LoadDSComms(activeDoc,"path")
    
	' File must be named "DS". Add code to let the user input a different to the excel file
	' the first time and save it for future calls of the routine (VERIFICAR SI SE HIZO BIEN Y BORRAR)
	
    Dim pathSufixDS As String = LoadDSComms(activeDoc,"document")
	Dim workSheetDS As String = LoadDSComms(activeDoc,"worksheet")
	
	Dim sFile As String = ruta & pathSufixDS
    Dim tabDS As String = workSheetDS
	
	Dim outputForDS(1) As String
	outputForDS(0) = sFile
	outputForDS(1) = tabDS
	
	Return outputForDS
	
	
End Function