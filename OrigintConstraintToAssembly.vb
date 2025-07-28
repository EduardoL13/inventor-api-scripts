Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document
AssemblyOriginCons(currentDoc) 
	
End Sub


Sub AssemblyOriginCons(assemDef As AssemblyDocument)

'Assembly origin planes declaration
Dim PlaneE1 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(1)
Dim PlaneE2 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(2)
Dim PlaneE3 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(3)
	
'For loop declaration and constraint imposition for each occurrence plane
For Each compOcc As ComponentOccurrence In assemDef.ComponentDefinition.Occurrences
	If compOcc.Constraints.Count = 0 Then
	
	' Occurrence origin planes declaration
	
	    Dim Plane1 As WorkPlane = compOcc.Definition.Workplanes.Item(1) 
	    Dim Plane2 As WorkPlane = compOcc.Definition.Workplanes.Item(2) 
	    Dim Plane3 As WorkPlane = compOcc.Definition.Workplanes.Item(3) 
		
	' Geometry proxy creation prior to constraint imposition
	
	    Dim APlane1 As WorkPlaneProxy
        compOcc.CreateGeometryProxy(Plane1,APlane1)
	    Dim APlane2 As WorkPlaneProxy
        compOcc.CreateGeometryProxy(Plane2,APlane2)
	    Dim APlane3 As WorkPlaneProxy
        compOcc.CreateGeometryProxy(Plane3, APlane3)
		
	' Origin planes constraint imposition
	
	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlane1,PlaneE1,0)
	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlane2,PlaneE2,0)
	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlane3,PlaneE3,0)
	
    End If
		
Next

End Sub	