Sub Main () ' v1
'Objetivo: escribir en un DS uno o más parámetros que se requieran de Inventor
' Declarations
Dim invDoc As PartDocument = ThisDoc.Document ' Documento activo

Dim inventorParamList As UserParameters = invDoc.ComponentDefinition.Parameters.UserParameters
Dim noPDS As Integer = inventorParamList.Count ' Insertar número de parámetros que se desean
Dim noListPDS As Integer = noPDS - 1 'Número de parámetros para poner en los arrays


Dim file As String
Dim tab As String
Dim outputForDS = getDS(invDoc)
file = outputForDS(0)
tab = "EIKP"

'Row counter definition
Dim RowCounter As Integer = 2 'Row de inicio para escribir datos


For Each param As UserParameter In inventorParamList
	If param.IsKey = True Then
        If RowCounter = 2 Then
	        GoExcel.CellValue(file, tab, "A" & RowCounter) = param.Name
		    cf = unitsEval(param.Units)
	        GoExcel.CellValue("B" & RowCounter) = param.Value * cf
	        GoExcel.CellValue("C" & RowCounter) = param.Units
    	Else
	        GoExcel.CellValue("A" & RowCounter) = param.Name
			'MsgBox(param.Name)
		    cf = unitsEval(param.Units)
	        GoExcel.CellValue("B" & RowCounter) = param.Value * cf
	        GoExcel.CellValue("C" & RowCounter) = param.Units
		End If
	RowCounter = RowCounter + 1 
	End If
Next

GoExcel.Save
MsgBox("Key parameters have been exported")
End Sub


Function unitsEval(units As String)
If units = "in"
	cf = 1 / 2.54
Else If units = "ul"
	cf = 1
Else
	cf = 10
End If
Return cf
End Function




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
	Else
		MsgBox("Input inválido")
	End If
End Function

Function getDS(activeDoc As Object)
	
	Dim ruta As String = LoadDSComms(activeDoc,"path")
    'MsgBox(ruta)
	
    Dim pathSufixDS As String = "DS.xlsx"
	'Dim workSheetDS As String = LoadDSComms(activeDoc,"worksheet")
	
	Dim sFile As String = ruta & pathSufixDS
    Dim tabDS As String = "EIKP"
	
	Dim outputForDS(1) As String
	outputForDS(0) = sFile
	outputForDS(1) = tabDS
	
	Return outputForDS
	
	
End Function

